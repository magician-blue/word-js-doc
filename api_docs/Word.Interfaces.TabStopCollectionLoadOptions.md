# Word.Interfaces.TabStopCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [tab stops](/en-us/javascript/api/word/word.tabstop) in a Word document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- alignment  
  For EACH ITEM in the collection: Gets a TabAlignment value that represents the alignment for the tab stop.

- customTab  
  For EACH ITEM in the collection: Gets whether this tab stop is a custom tab stop.

- leader  
  For EACH ITEM in the collection: Gets a TabLeader value that represents the leader for this TabStop object.

- next  
  For EACH ITEM in the collection: Gets the next tab stop in the collection.

- position  
  For EACH ITEM in the collection: Gets the position of the tab stop relative to the left margin.

- previous  
  For EACH ITEM in the collection: Gets the previous tab stop in the collection.

## Property Details

### $all

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value  
boolean

---

### alignment

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a TabAlignment value that represents the alignment for the tab stop.

```typescript
alignment?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### customTab

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets whether this tab stop is a custom tab stop.

```typescript
customTab?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### leader

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a TabLeader value that represents the leader for this TabStop object.

```typescript
leader?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### next

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the next tab stop in the collection.

```typescript
next?: Word.Interfaces.TabStopLoadOptions;
```

Property Value  
[Word.Interfaces.TabStopLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstoploadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### position

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the position of the tab stop relative to the left margin.

```typescript
position?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### previous

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the previous tab stop in the collection.

```typescript
previous?: Word.Interfaces.TabStopLoadOptions;
```

Property Value  
[Word.Interfaces.TabStopLoadOptions](/en-us/javascript/api/word/word.interfaces.tabstoploadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]