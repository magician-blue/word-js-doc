# Word.DocumentPropertyType enum

Package: [word](/en-us/javascript/api/word)

## Remarks

[ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

## Fields

- boolean = "Boolean"
  - [ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- date = "Date"
  - [ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- number = "Number"
  - [ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- string = "String"
  - [ API set: WordApi 1.3 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)