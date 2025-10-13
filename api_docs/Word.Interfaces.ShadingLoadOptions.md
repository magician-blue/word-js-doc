# Word.Interfaces.ShadingLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the shading object.

## Remarks

[ API set: WordApi 1.6 ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- backgroundPatternColor  
  Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

- foregroundPatternColor  
  Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

- texture  
  Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word (https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value  
boolean

---

### backgroundPatternColor

Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundPatternColor?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi 1.6 ]

---

### foregroundPatternColor

Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundPatternColor?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApiDesktop 1.1 ]

---

### texture

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word (https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

```typescript
texture?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApiDesktop 1.1 ]