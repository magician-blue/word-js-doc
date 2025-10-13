# Word.ListTemplateApplyOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents options for applying a list template to a range.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- applyLevel  
  If provided, specifies the level to apply in the list template. The default value is 1.

- applyTo  
  If provided, specifies which part of the list to apply the template to. The default value is `Word.ListApplyTo.wholeList`.

- continuePreviousList  
  If provided, specifies whether to continue the previous list. The default value is `false`.

- defaultListBehavior  
  If provided, specifies the default list behavior. The default value is `DefaultListBehavior.word97`.

## Property Details

### applyLevel

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the level to apply in the list template. The default value is 1.

```typescript
applyLevel?: number;
```

Property Value
number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### applyTo

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies which part of the list to apply the template to. The default value is `Word.ListApplyTo.wholeList`.

```typescript
applyTo?: Word.ListApplyTo | "WholeList" | "ThisPointForward" | "Selection";
```

Property Value  
[Word.ListApplyTo](/en-us/javascript/api/word/word.listapplyto) | "WholeList" | "ThisPointForward" | "Selection"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### continuePreviousList

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to continue the previous list. The default value is `false`.

```typescript
continuePreviousList?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### defaultListBehavior

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the default list behavior. The default value is `DefaultListBehavior.word97`.

```typescript
defaultListBehavior?: Word.DefaultListBehavior | "Word97" | "Word2000" | "Word2002";
```

Property Value  
[Word.DefaultListBehavior](/en-us/javascript/api/word/word.defaultlistbehavior) | "Word97" | "Word2000" | "Word2002"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)