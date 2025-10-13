# Word.CustomXmlAddSchemaOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds one or more schemas to a schema collection that can then be added to a stream in the data store and to the schema library.

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- alias: If provided, specifies the alias of the schema to be added to the collection. However, if the alias already exists in the Schema Library, the schema can be found using this value.
- fileName: If provided, specifies the location of the schema on a disk. If this property is specified, the schema is added to the collection and to the Schema Library.
- installForAllUsers: If provided, specifies whether, in the case where the schema is being added to the Schema Library, the Schema Library keys should be written to the registry (`HKEY_LOCAL_MACHINE` for all users or `HKEY_CURRENT_USER` for just the current user). The property defaults to `false` and writes to `HKEY_CURRENT_USER`.
- namespaceUri: If provided, specifies the namespace of the schema to be added to the collection. However, if the schema already exists in the Schema Library, the schema will be retrieved from there.

## Property Details

### alias

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the alias of the schema to be added to the collection. However, if the alias already exists in the Schema Library, the schema can be found using this value.

```typescript
alias?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### fileName

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the location of the schema on a disk. If this property is specified, the schema is added to the collection and to the Schema Library.

```typescript
fileName?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### installForAllUsers

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether, in the case where the schema is being added to the Schema Library, the Schema Library keys should be written to the registry (`HKEY_LOCAL_MACHINE` for all users or `HKEY_CURRENT_USER` for just the current user). The property defaults to `false` and writes to `HKEY_CURRENT_USER`.

```typescript
installForAllUsers?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### namespaceUri

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the namespace of the schema to be added to the collection. However, if the schema already exists in the Schema Library, the schema will be retrieved from there.

```typescript
namespaceUri?: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)