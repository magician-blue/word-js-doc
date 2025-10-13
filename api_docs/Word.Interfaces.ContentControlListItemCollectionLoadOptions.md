# Word.Interfaces.ContentControlListItemCollectionLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem) objects that represent the items in a dropdown list or combo box content control.

## Remarks

[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  - Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- displayText  
  - For EACH ITEM in the collection: Specifies the display text of a list item for a dropdown list or combo box content control.

- index  
  - For EACH ITEM in the collection: Specifies the index location of a content control list item in the collection of list items.

- value  
  - For EACH ITEM in the collection: Specifies the programmatic value of a list item for a dropdown list or combo box content control.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property Value: boolean

### displayText

For EACH ITEM in the collection: Specifies the display text of a list item for a dropdown list or combo box content control.

```typescript
displayText?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

For EACH ITEM in the collection: Specifies the index location of a content control list item in the collection of list items.

```typescript
index?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value

For EACH ITEM in the collection: Specifies the programmatic value of a list item for a dropdown list or combo box content control.

```typescript
value?: boolean;
```

- Property Value: boolean  
- Remarks: [API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)