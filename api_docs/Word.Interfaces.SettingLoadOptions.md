# Word.Interfaces.SettingLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a setting of the add-in.

## Remarks

[ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- key  
  Gets the key of the setting.

- value  
  Specifies the value of the setting.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property Value: boolean

### key

Gets the key of the setting.

```typescript
key?: boolean;
```

- Property Value: boolean  
- Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value

Specifies the value of the setting.

```typescript
value?: boolean;
```

- Property Value: boolean  
- Remarks: [ API set: WordApi 1.4 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)