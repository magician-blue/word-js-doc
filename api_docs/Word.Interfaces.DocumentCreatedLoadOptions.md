# Word.Interfaces.DocumentCreatedLoadOptions interface

Package: [word](/en-us/javascript/api/word)

The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.

## Remarks

[API set: WordApi 1.3]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- body  
  Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

- properties  
  Gets the properties of the document.

- saved  
  Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### body

Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[API set: WordApiHiddenDocument 1.3]

---

### properties

Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesLoadOptions;
```

Property value: [Word.Interfaces.DocumentPropertiesLoadOptions](/en-us/javascript/api/word/word.interfaces.documentpropertiesloadoptions)

Remarks  
[API set: WordApiHiddenDocument 1.3]

---

### saved

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
saved?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApiHiddenDocument 1.3]