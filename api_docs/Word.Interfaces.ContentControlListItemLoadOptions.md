# Word.Interfaces.ContentControlListItemLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a list item in a dropdown list or combo box content control.

## Remarks

[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- displayText  
  Specifies the display text of a list item for a dropdown list or combo box content control.

- index  
  Specifies the index location of a content control list item in the collection of list items.

- value  
  Specifies the programmatic value of a list item for a dropdown list or combo box content control.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### displayText

Specifies the display text of a list item for a dropdown list or combo box content control.

```typescript
displayText?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### index

Specifies the index location of a content control list item in the collection of list items.

```typescript
index?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### value

Specifies the programmatic value of a list item for a dropdown list or combo box content control.

```typescript
value?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)