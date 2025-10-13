# Word.Interfaces.ListTemplateData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling listTemplate.toJSON().

## Properties

- listLevels  
  Gets a ListLevelCollection object that represents all the levels for the list template.
- outlineNumbered  
  Specifies whether the list template is outline numbered.

## Property Details

### listLevels

Gets a ListLevelCollection object that represents all the levels for the list template.

```typescript
listLevels?: Word.Interfaces.ListLevelData[];
```

Property Value  
[Word.Interfaces.ListLevelData](/en-us/javascript/api/word/word.interfaces.listleveldata)[]

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### outlineNumbered

Specifies whether the list template is outline numbered.

```typescript
outlineNumbered?: boolean;
```

Property Value  
boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)