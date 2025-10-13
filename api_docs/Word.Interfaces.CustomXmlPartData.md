# Word.Interfaces.CustomXmlPartData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `customXmlPart.toJSON()`.

## Properties

- builtIn  
  Gets a value that indicates whether the `CustomXmlPart` is built-in.

- documentElement  
  Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

- errors  
  Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.

- id  
  Gets the ID of the custom XML part.

- namespaceManager  
  Gets the set of namespace prefix mappings used against the current `CustomXmlPart` object.

- namespaceUri  
  Gets the namespace URI of the custom XML part.

- schemaCollection  
  Specifies a `CustomXmlSchemaCollection` object representing the set of schemas attached to a bound region of data in the document.

- xml  
  Gets the XML representation of the current `CustomXmlPart` object.

## Property Details

### builtIn

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a value that indicates whether the `CustomXmlPart` is built-in.

```typescript
builtIn?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### documentElement

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

```typescript
documentElement?: Word.Interfaces.CustomXmlNodeData;
```

Property Value
- [Word.Interfaces.CustomXmlNodeData](/en-us/javascript/api/word/word.interfaces.customxmlnodedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### errors

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `CustomXmlValidationErrorCollection` object that provides access to any XML validation errors.

```typescript
errors?: Word.Interfaces.CustomXmlValidationErrorData[];
```

Property Value
- [Word.Interfaces.CustomXmlValidationErrorData](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrordata)[]

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets the ID of the custom XML part.

```typescript
id?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### namespaceManager

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the set of namespace prefix mappings used against the current `CustomXmlPart` object.

```typescript
namespaceManager?: Word.Interfaces.CustomXmlPrefixMappingData[];
```

Property Value
- [Word.Interfaces.CustomXmlPrefixMappingData](/en-us/javascript/api/word/word.interfaces.customxmlprefixmappingdata)[]

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### namespaceUri

Gets the namespace URI of the custom XML part.

```typescript
namespaceUri?: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### schemaCollection

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `CustomXmlSchemaCollection` object representing the set of schemas attached to a bound region of data in the document.

```typescript
schemaCollection?: Word.Interfaces.CustomXmlSchemaData[];
```

Property Value
- [Word.Interfaces.CustomXmlSchemaData](/en-us/javascript/api/word/word.interfaces.customxmlschemadata)[]

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xml

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the XML representation of the current `CustomXmlPart` object.

```typescript
xml?: string;
```

Property Value
- string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)