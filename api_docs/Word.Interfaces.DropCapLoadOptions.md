# Word.Interfaces.DropCapLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a dropped capital letter in a Word document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- distanceFromText  
  Gets the distance (in points) between the dropped capital letter and the paragraph text.

- fontName  
  Gets the name of the font for the dropped capital letter.

- linesToDrop  
  Gets the height (in lines) of the dropped capital letter.

- position  
  Gets the position of the dropped capital letter.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

---

### distanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the distance (in points) between the dropped capital letter and the paragraph text.

```typescript
distanceFromText?: boolean;
```

Property value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### fontName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the font for the dropped capital letter.

```typescript
fontName?: boolean;
```

Property value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### linesToDrop

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in lines) of the dropped capital letter.

```typescript
linesToDrop?: boolean;
```

Property value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### position

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the dropped capital letter.

```typescript
position?: boolean;
```

Property value: boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]