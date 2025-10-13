# Word.WindowScrollOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options that scrolls a window or pane by the specified number of units defined by the calling method.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- down  
  If provided, specifies the number of units to scroll the window down. If `down` and `up` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `down` is 3 and `up` is 6, the contents are scrolled up three units.

- left  
  If provided, specifies the number of screens to scroll the window to the left. If `left` and `right` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `left` is 3 and `right` is 6, the contents are scrolled to the right three screens.

- right  
  If provided, specifies the number of screens to scroll the window to the right. If `left` and `right` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `left` is 3 and `right` is 6, the contents are scrolled to the right three screens.

- up  
  If provided, specifies the number of units to scroll the window up. If `down` and `up` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `down` is 3 and `up` is 6, the contents are scrolled up three units.

## Property Details

### down

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of units to scroll the window down. If `down` and `up` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `down` is 3 and `up` is 6, the contents are scrolled up three units.

```typescript
down?: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### left

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of screens to scroll the window to the left. If `left` and `right` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `left` is 3 and `right` is 6, the contents are scrolled to the right three screens.

```typescript
left?: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### right

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of screens to scroll the window to the right. If `left` and `right` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `left` is 3 and `right` is 6, the contents are scrolled to the right three screens.

```typescript
right?: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### up

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of units to scroll the window up. If `down` and `up` are both provided, the contents of the window are scrolled by the difference of the property values. For example, if `down` is 3 and `up` is 6, the contents are scrolled up three units.

```typescript
up?: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)