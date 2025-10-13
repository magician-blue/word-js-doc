# Word.Interfaces.SettingCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains the collection of [Word.Setting](/en-us/javascript/api/word/word.setting) objects.

## Remarks

[ API set: WordApi 1.4 ]

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- key: For EACH ITEM in the collection: Gets the key of the setting.
- value: For EACH ITEM in the collection: Specifies the value of the setting.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### key

For EACH ITEM in the collection: Gets the key of the setting.

```typescript
key?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### value

For EACH ITEM in the collection: Specifies the value of the setting.

```typescript
value?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]