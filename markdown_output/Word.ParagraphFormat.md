# Word.ParagraphFormat

**Package:** `word`

**API Set:** WordApi 1.5

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a style of paragraph in a document.

## Class Examples

```typescript
// Link to full sample: // Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

## Properties

### alignment

**Type:** `Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"`

**Since:** WordApi 1.5

Specifies the alignment for the specified paragraphs.

#### Examples

**Example**: Update the paragraph format of a specified style by setting its left indent to 30 and alignment to centered.

```typescript
// Link to full sample: // Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ParagraphFormat object to verify the connection to the Office host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const paragraphFormat = paragraph.paragraphFormat;
    
    // Load the paragraph format
    paragraphFormat.load("alignment");
    await context.sync();
    
    // Access the request context from the ParagraphFormat object
    const requestContext = paragraphFormat.context;
    
    // Verify the context is available and log information
    console.log("Request context is connected:", requestContext !== null);
    console.log("Context debug info:", requestContext.debugInfo);
    
    await context.sync();
});
```

---

### firstLineIndent

**Type:** `number`

**Since:** WordApi 1.5

Specifies the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

#### Examples

**Example**: Set a first-line indent of 36 points for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.format.firstLineIndent = 36;
    
    await context.sync();
});
```

---

### keepTogether

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.

#### Examples

**Example**: Keep all lines of the first paragraph together on the same page to prevent it from being split across pages during repagination

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.format.keepTogether = true;
    
    await context.sync();
});
```

---

### keepWithNext

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.

#### Examples

**Example**: Keep the selected paragraph together with the next paragraph when Word repaginates the document to prevent them from being split across pages.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.format.keepWithNext = true;
    
    await context.sync();
    console.log("Paragraph will stay with the next paragraph during pagination.");
});
```

---

### leftIndent

**Type:** `number`

**Since:** WordApi 1.5

Specifies the left indent.

#### Examples

**Example**: Modify a style's paragraph format by setting the left indent to 30 points and the alignment to centered.

```typescript
// Link to full sample: // Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

---

### lineSpacing

**Type:** `number`

**Since:** WordApi 1.5

Specifies the line spacing (in points) for the specified paragraphs.

#### Examples

**Example**: Set the line spacing to 18 points for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.paragraphFormat.lineSpacing = 18;
    
    await context.sync();
});
```

---

### lineUnitAfter

**Type:** `number`

**Since:** WordApi 1.5

Specifies the amount of spacing (in gridlines) after the specified paragraphs.

#### Examples

**Example**: Set the spacing after a paragraph to 2 gridlines to add vertical space between paragraphs in a document

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.format.lineUnitAfter = 2;
    
    await context.sync();
});
```

---

### lineUnitBefore

**Type:** `number`

**Since:** WordApi 1.5

Specifies the amount of spacing (in gridlines) before the specified paragraphs.

#### Examples

**Example**: Set the spacing before a paragraph to 2 gridlines to create vertical space above it

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.format.lineUnitBefore = 2;
    
    await context.sync();
});
```

---

### mirrorIndents

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether left and right indents are the same width.

#### Examples

**Example**: Set mirror indents on the first paragraph so that left and right indents have the same width

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.paragraphFormat.mirrorIndents = true;
    
    await context.sync();
});
```

---

### outlineLevel

**Type:** `Word.OutlineLevel | "OutlineLevel1" | "OutlineLevel2" | "OutlineLevel3" | "OutlineLevel4" | "OutlineLevel5" | "OutlineLevel6" | "OutlineLevel7" | "OutlineLevel8" | "OutlineLevel9" | "OutlineLevelBodyText"`

**Since:** WordApi 1.5

Specifies the outline level for the specified paragraphs.

#### Examples

**Example**: Set the outline level of the first paragraph to Level 2 to organize document structure for navigation and table of contents.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.format.outlineLevel = Word.OutlineLevel.level2;
    
    await context.sync();
});
```

---

### rightIndent

**Type:** `number`

**Since:** WordApi 1.5

Specifies the right indent (in points) for the specified paragraphs.

#### Examples

**Example**: Set the right indent of the first paragraph to 36 points (0.5 inches) to create a narrower text margin on the right side.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.paragraphFormat.rightIndent = 36;
    
    await context.sync();
});
```

---

### spaceAfter

**Type:** `number`

**Since:** WordApi 1.5

Specifies the amount of spacing (in points) after the specified paragraph or text column.

#### Examples

**Example**: Set the spacing after a paragraph to 12 points to add vertical space between paragraphs

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.format.spaceAfter = 12;
    
    await context.sync();
});
```

---

### spaceBefore

**Type:** `number`

**Since:** WordApi 1.5

Specifies the spacing (in points) before the specified paragraphs.

#### Examples

**Example**: Set the spacing before a paragraph to 12 points to add space above it

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.format.spaceBefore = 12;
    
    await context.sync();
});
```

---

### widowControl

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Microsoft Word repaginates the document.

#### Examples

**Example**: Disable widow control for the first paragraph to allow its first and last lines to appear on different pages when Word repaginates the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.paragraphFormat.widowControl = false;
    
    await context.sync();
    console.log("Widow control disabled for the first paragraph");
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ParagraphFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ParagraphFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ParagraphFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ParagraphFormat`

#### Examples

**Example**: Read and display the alignment and indentation properties of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const paragraphFormat = paragraph.paragraphFormat;
    
    // Load specific properties of the paragraph format
    paragraphFormat.load("alignment, leftIndent, rightIndent, firstLineIndent");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Alignment:", paragraphFormat.alignment);
    console.log("Left Indent:", paragraphFormat.leftIndent);
    console.log("Right Indent:", paragraphFormat.rightIndent);
    console.log("First Line Indent:", paragraphFormat.firstLineIndent);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ParagraphFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ParagraphFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Format the first paragraph by setting multiple properties at once including alignment, indentation, and spacing

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    firstParagraph.format.set({
        alignment: Word.Alignment.centered,
        firstLineIndent: 36,
        leftIndent: 72,
        rightIndent: 72,
        spaceAfter: 12,
        spaceBefore: 6
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ParagraphFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ParagraphFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ParagraphFormatData`

#### Examples

**Example**: Serialize paragraph format properties to JSON for logging or storage purposes

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const format = paragraph.paragraphFormat;
    
    // Load the paragraph format properties
    format.load("alignment,firstLineIndent,leftIndent,rightIndent,spaceAfter,spaceBefore");
    
    await context.sync();
    
    // Convert the paragraph format to a plain JSON object
    const formatData = format.toJSON();
    
    // Log the serialized format data
    console.log("Paragraph Format as JSON:", JSON.stringify(formatData, null, 2));
    
    // The JSON object can now be stored, transmitted, or compared
    // Example output: { alignment: "Left", firstLineIndent: 0, leftIndent: 0, ... }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ParagraphFormat`

#### Examples

**Example**: Track a paragraph's format object across multiple sync calls to monitor and preserve its alignment changes without encountering InvalidObjectPath errors.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const paragraphFormat = paragraph.paragraphFormat;
    
    // Track the format object to use it across multiple sync calls
    paragraphFormat.track();
    
    // Load initial properties
    paragraphFormat.load("alignment");
    await context.sync();
    
    console.log("Initial alignment:", paragraphFormat.alignment);
    
    // Change alignment
    paragraphFormat.alignment = Word.Alignment.centered;
    await context.sync();
    
    // Access the tracked object again after sync without errors
    paragraphFormat.load("alignment");
    await context.sync();
    
    console.log("Updated alignment:", paragraphFormat.alignment);
    
    // Untrack when done
    paragraphFormat.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ParagraphFormat`

#### Examples

**Example**: Format a paragraph and then release it from memory tracking to optimize performance after the formatting is complete.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const paragraphFormat = paragraph.paragraphFormat;
    
    // Track the paragraph format object for changes
    paragraphFormat.load("alignment");
    context.trackedObjects.add(paragraphFormat);
    
    await context.sync();
    
    // Make changes to the paragraph format
    paragraphFormat.alignment = Word.Alignment.centered;
    paragraphFormat.leftIndent = 20;
    paragraphFormat.spaceAfter = 12;
    
    await context.sync();
    
    // Release the tracked object from memory after we're done using it
    paragraphFormat.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://docs.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml
