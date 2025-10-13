# Word.WindowCloseOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options that define whether to save changes before closing and whether to route the document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- routeDocument  
  If provided, specifies whether to route the document to the next recipient. If the document doesn't have a routing slip attached, this property is ignored.

- saveChanges  
  If provided, specifies the save action for the document. For available values, see [Word.SaveConfiguration](/en-us/javascript/api/word/word.saveconfiguration).

## Property Details

### routeDocument

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to route the document to the next recipient. If the document doesn't have a routing slip attached, this property is ignored.

```typescript
routeDocument?: boolean;
```

Property Value

- boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### saveChanges

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the save action for the document. For available values, see [Word.SaveConfiguration](/en-us/javascript/api/word/word.saveconfiguration).

```typescript
saveChanges?: Word.SaveConfiguration | "DoNotSaveChanges" | "SaveChanges" | "PromptToSaveChanges";
```

Property Value

- [Word.SaveConfiguration](/en-us/javascript/api/word/word.saveconfiguration) | "DoNotSaveChanges" | "SaveChanges" | "PromptToSaveChanges"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]