# Word.Interfaces.TabStopLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a tab stop in a Word document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- alignment  
  Gets a `TabAlignment` value that represents the alignment for the tab stop.

- customTab  
  Gets whether this tab stop is a custom tab stop.

- leader  
  Gets a `TabLeader` value that represents the leader for this `TabStop` object.

- next  
  Gets the next tab stop in the collection.

- position  
  Gets the position of the tab stop relative to the left margin.

- previous  
  Gets the previous tab stop in the collection.

## Property Details

### $all

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### alignment

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TabAlignment` value that represents the alignment for the tab stop.

```typescript
alignment?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### customTab

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether this tab stop is a custom tab stop.

```typescript
customTab?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leader

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TabLeader` value that represents the leader for this `TabStop` object.

```typescript
leader?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### next

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next tab stop in the collection.

```typescript
next?: Word.Interfaces.TabStopLoadOptions;
```

Property Value: [Word.Interfaces.TabStopLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstoploadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### position

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the tab stop relative to the left margin.

```typescript
position?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### previous

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous tab stop in the collection.

```typescript
previous?: Word.Interfaces.TabStopLoadOptions;
```

Property Value: [Word.Interfaces.TabStopLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstoploadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)