# RequestContext

**Package:** `word`

**API Set:** None None

**Extends:** `OfficeExtension.ClientRequestContext`

## Description

The RequestContext object facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the request context is required to get access to the Word object model from the add-in.

## Class Examples

```typescript
// *.run methods automatically create an OfficeExtension.ClientRequestContext
// object to work with the Office file.
await Word.run(async (context: Word.RequestContext) => {
  const document = context.document;
  // Interact with the Word document...
});
```

## Properties

### application

**Type:** `Application`

**Since:** WordApi 1.3

#### Examples

**Example**: Get the application name and version information from the Word application

```typescript
await Word.run(async (context: Word.RequestContext) => {
    const app = context.application;
    
    // Load the application properties
    app.load("name, version");
    
    await context.sync();
    
    console.log(`Application: ${app.name}`);
    console.log(`Version: ${app.version}`);
});
```

---

### document

**Type:** `Document`

#### Examples

**Example**: Access the document body and insert text at the beginning of the Word document

```typescript
await Word.run(async (context: Word.RequestContext) => {
    // Access the document through the context
    const document = context.document;
    const body = document.body;
    
    // Insert text at the start of the document
    body.insertText("Hello from Word.js API!", Word.InsertLocation.start);
    
    await context.sync();
});
```

---

## Methods

### constructor

**Kind:** `create`

Constructs a new instance of the RequestContext class

#### Signature

**Parameters:**
- `url`: `string` (optional)

**Returns:** `None`

#### Examples

**Example**: Create a request context for a specific Word document using its URL to enable cross-document operations

```typescript
async function accessSpecificDocument() {
    // URL of another Word document open in the same session
    const documentUrl = "https://contoso.sharepoint.com/shared/Document.docx";
    
    // Create a new RequestContext for the specific document
    const context = new Word.RequestContext(documentUrl);
    
    await context.sync();
    
    // Now you can work with the specified document
    const body = context.document.body;
    body.insertText("Text inserted via cross-document context", Word.InsertLocation.end);
    
    await context.sync();
}
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
