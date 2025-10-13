# Word.Interfaces.DocumentLibraryVersionData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `documentLibraryVersion.toJSON()`.

## Properties

- comments: Gets any optional comments associated with this version of the shared document.
- modified: Gets the date and time at which this version of the shared document was last saved to the server.
- modifiedBy: Gets the name of the user who last saved this version of the shared document to the server.

## Property Details

### comments

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets any optional comments associated with this version of the shared document.

```typescript
comments?: string;
```

Property value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### modified

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date and time at which this version of the shared document was last saved to the server.

```typescript
modified?: any;
```

Property value: any

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### modifiedBy

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the user who last saved this version of the shared document to the server.

```typescript
modifiedBy?: string;
```

Property value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)