# Word.RequestContext class

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Extends: [OfficeExtension.ClientRequestContext](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext)

The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.

## Remarks

#### Examples

```typescript
// *.run methods automatically create an OfficeExtension.ClientRequestContext
// object to work with the Office file.
await Word.run(async (context: Word.RequestContext) => {
  const document = context.document;
  // Interact with the Word document...
});
```

## Constructors

- (constructor)(url) — Constructs a new instance of the RequestContext class

## Properties

- application — [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) *
- document

## Constructor Details

### (constructor)(url)

Constructs a new instance of the RequestContext class

Signature:
```typescript
constructor(url?: string);
```

Parameters:
- url: string

## Property Details

### application

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) *

```typescript
readonly application: Application;
```

Property Value:
- [Word.Application](https://learn.microsoft.com/en-us/javascript/api/word/word.application)

### document

```typescript
readonly document: Document;
```

Property Value:
- [Word.Document](https://learn.microsoft.com/en-us/javascript/api/word/word.document)