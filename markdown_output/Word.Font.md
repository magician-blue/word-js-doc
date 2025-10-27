# Word.Font

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApi 1.1

**Extends:** `https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject`

## Description

Represents a font.

## Class Examples

```typescript
// Change the font color
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to change the font color of the current selection.
    selection.font.color = 'blue';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The font color of the selection has been changed.');
});
```

## Properties

### allCaps

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
- true: All the text has the All Caps attribute.
- false: None of the text has the All Caps attribute.
- null: Returned if some, but not all, of the text has the All Caps attribute.

#### Examples

**Example**: Set the selected text to display in all capital letters using the allCaps formatting

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.allCaps = true;
    
    await context.sync();
});
```

---

### bold

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

#### Examples

**Example**: Apply bold formatting to the currently selected text in the document.

```typescript
// Bold format text
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to make the current selection bold.
    selection.font.bold = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The selection is now bold.');
});
```

---

### boldBidirectional

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
- true: All the text is bold.
- false: None of the text is bold.
- null: Returned if some, but not all, of the text is bold.

#### Examples

**Example**: Set the selected text to bold formatting for right-to-left language content (such as Arabic or Hebrew)

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Set bold formatting for right-to-left text
    font.boldBidirectional = true;
    
    await context.sync();
});
```

---

### borders

**Type:** `Word.BorderUniversalCollection`

**Since:** WordApi BETA

Returns a BorderUniversalCollection object that represents all the borders for the font.

#### Examples

**Example**: Add a red double-line border around all text formatted with a specific font style

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's font
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Access the borders collection for the font
    const borders = font.borders;
    borders.load("items");
    
    await context.sync();
    
    // Set border properties for all borders in the collection
    borders.items.forEach(border => {
        border.type = Word.BorderType.double;
        border.color = "#FF0000"; // Red color
        border.width = 2;
    });
    
    await context.sync();
});
```

---

### color

**Type:** `string`

**Since:** WordApi 1.1

Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

#### Examples

**Example**: Change the font color of the currently selected text to blue.

```typescript
// Change the font color
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to change the font color of the current selection.
    selection.font.color = 'blue';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The font color of the selection has been changed.');
});
```

---

### colorIndex

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

**Since:** WordApi BETA

Specifies a ColorIndex value that represents the color for the font.

#### Examples

**Example**: Set the font color of the selected text to red using the color index

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.colorIndex = "Red";
    
    await context.sync();
});
```

---

### colorIndexBidirectional

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

**Since:** WordApi BETA

Specifies the color for the Font object in a right-to-left language document.

#### Examples

**Example**: Set the font color to red for Arabic text in a right-to-left language document

```typescript
await Word.run(async (context) => {
    // Get the selected text
    const range = context.document.getSelection();
    
    // Set the color index for right-to-left text to red
    range.font.colorIndexBidirectional = "Red";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context to verify the Office host connection before applying font formatting operations

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Access the request context associated with the font object
    const fontContext = font.context;
    
    // Use the context to verify connection and perform operations
    console.log("Context connected:", fontContext !== null);
    
    // Apply font formatting using the same context
    font.bold = true;
    font.size = 14;
    
    await context.sync();
});
```

---

### contextualAlternates

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether contextual alternates are enabled for the font.

#### Examples

**Example**: Enable contextual alternates for the selected text to improve the appearance of letter combinations

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.contextualAlternates = true;
    
    await context.sync();
});
```

---

### diacriticColor

**Type:** `string`

**Since:** WordApi BETA

Specifies the color to be used for diacritics for the Font object. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the diacritics color to red for the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Set diacritic color to red
    font.diacriticColor = "#FF0000";
    
    await context.sync();
});
```

---

### disableCharacterSpaceGrid

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether Microsoft Word ignores the number of characters per line for the corresponding Font object.

#### Examples

**Example**: Disable the character space grid for the selected text so that Word ignores character-per-line spacing constraints

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.disableCharacterSpaceGrid = true;
    
    await context.sync();
});
```

---

### doubleStrikeThrough

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

#### Examples

**Example**: Apply double strikethrough formatting to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Apply double strikethrough to the selected text
    font.doubleStrikeThrough = true;
    
    await context.sync();
});
```

---

### emboss

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is formatted as embossed. The possible values are as follows:
- true: All the text is embossed.
- false: None of the text is embossed.
- null: Returned if some, but not all, of the text is embossed.

#### Examples

**Example**: Apply embossed formatting to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Apply embossed formatting to the selected text
    font.emboss = true;
    
    await context.sync();
});
```

---

### emphasisMark

**Type:** `Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"`

**Since:** WordApi BETA

Specifies an EmphasisMark value that represents the emphasis mark for a character or designated character string.

#### Examples

**Example**: Add emphasis marks above the selected text using solid circles to highlight important terms in a document.

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Set emphasis mark to display solid circles over the text
    font.emphasisMark = Word.EmphasisMark.overSolidCircle;
    
    await context.sync();
});
```

---

### engrave

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is formatted as engraved. The possible values are as follows:
- true: All the text is engraved.
- false: None of the text is engraved.
- null: Returned if some, but not all, of the text is engraved.

#### Examples

**Example**: Apply engraved formatting to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Apply engraved formatting to the selected text
    font.engrave = true;
    
    await context.sync();
});
```

---

### fill

**Type:** `Word.FillFormat`

**Since:** WordApi BETA

Returns a FillFormat object that contains fill formatting properties for the font used by the range of text.

#### Examples

**Example**: Set the font fill color to solid blue for the selected text in the document.

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    const fill = font.fill;
    
    fill.setSolidColor("blue");
    
    await context.sync();
});
```

---

### glow

**Type:** `Word.GlowFormat`

**Since:** WordApi BETA

Returns a GlowFormat object that represents the glow formatting for the font used by the range of text.

#### Examples

**Example**: Apply a blue glow effect with 8pt size to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    const glow = font.glow;
    
    glow.color = "blue";
    glow.radius = 8;
    glow.transparency = 0.5;
    
    await context.sync();
});
```

---

### hidden

**Type:** `boolean`

**Since:** WordApiDesktop 1.2

Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

#### Examples

**Example**: Mark selected text as hidden so it won't be visible when printed or displayed

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.hidden = true;
    
    await context.sync();
});
```

---

### highlightColor

**Type:** `string`

**Since:** WordApi 1.1

Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

#### Examples

**Example**: Highlight the currently selected text in the document with a yellow background color.

```typescript
// Highlight selected text
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The selection has been highlighted.');
});
```

---

### italic

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

#### Examples

**Example**: Make the text in the first paragraph italic

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.font.italic = true;
    
    await context.sync();
});
```

---

### italicBidirectional

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
- true: All the text is italicized.
- false: None of the text is italicized.
- null: Returned if some, but not all, of the text is italicized.

#### Examples

**Example**: Set the selected text to italic formatting for right-to-left language content (such as Arabic or Hebrew)

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Set italic formatting for right-to-left text
    font.italicBidirectional = true;
    
    await context.sync();
});
```

---

### kerning

**Type:** `number`

**Since:** WordApi BETA

Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

#### Examples

**Example**: Set the minimum font size to 12 points for automatic kerning adjustment on the selected text

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.kerning = 12;
    
    await context.sync();
});
```

---

### ligature

**Type:** `Word.Ligature | "None" | "Standard" | "Contextual" | "StandardContextual" | "Historical" | "StandardHistorical" | "ContextualHistorical" | "StandardContextualHistorical" | "Discretional" | "StandardDiscretional" | "ContextualDiscretional" | "StandardContextualDiscretional" | "HistoricalDiscretional" | "StandardHistoricalDiscretional" | "ContextualHistoricalDiscretional" | "All"`

**Since:** WordApi BETA

Specifies the ligature setting for the Font object.

#### Examples

**Example**: Set the ligature setting to "Standard" for the selected text to enable standard typographic ligatures

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.ligature = "Standard";
    
    await context.sync();
});
```

---

### line

**Type:** `Word.LineFormat`

**Since:** WordApi BETA

Returns a LineFormat object that specifies the formatting for a line.

#### Examples

**Example**: Set the underline style to single and change the underline color to red for selected text

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    const line = font.line;
    
    line.underlineStyle = Word.UnderlineType.single;
    line.underlineColor = "red";
    
    await context.sync();
});
```

---

### name

**Type:** `string`

**Since:** WordApi 1.1

Specifies a value that represents the name of the font.

#### Examples

**Example**: Change the font name of the currently selected text to Arial.

```typescript
// Change the font name
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to change the current selection's font name.
    selection.font.name = 'Arial';

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The font name has changed.');
});
```

---

### nameAscii

**Type:** `string`

**Since:** WordApi BETA

Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

#### Examples

**Example**: Set the font for Latin characters in the selected text to "Courier New"

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.font.nameAscii = "Courier New";
    
    await context.sync();
});
```

---

### nameBidirectional

**Type:** `string`

**Since:** WordApi BETA

Specifies the font name in a right-to-left language document.

#### Examples

**Example**: Set the bidirectional font name to "Arial" for the selected text in a right-to-left language document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.nameBidirectional = "Arial";
    
    await context.sync();
});
```

---

### nameFarEast

**Type:** `string`

**Since:** WordApi BETA

Specifies the East Asian font name.

#### Examples

**Example**: Set the East Asian font name to "MS Mincho" for the selected text in the document.

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.font.nameFarEast = "MS Mincho";
    
    await context.sync();
});
```

---

### nameOther

**Type:** `string`

**Since:** WordApi BETA

Specifies the font used for characters with codes from 128 through 255.

#### Examples

**Example**: Set the font for extended ASCII characters (128-255) to "Courier New" in the selected text

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.nameOther = "Courier New";
    
    await context.sync();
});
```

---

### numberForm

**Type:** `Word.NumberForm | "Default" | "Lining" | "OldStyle"`

**Since:** WordApi BETA

Specifies the number form setting for an OpenType font.

#### Examples

**Example**: Set the number form to old-style figures for the selected text to give numbers a more traditional, varying baseline appearance

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.numberForm = Word.NumberForm.oldStyle;
    
    await context.sync();
});
```

---

### numberSpacing

**Type:** `Word.NumberSpacing | "Default" | "Proportional" | "Tabular"`

**Since:** WordApi BETA

Specifies the number spacing setting for the font.

#### Examples

**Example**: Set the number spacing of selected text to tabular format so that numbers align vertically in columns

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.numberSpacing = Word.NumberSpacing.tabular;
    
    await context.sync();
});
```

---

### outline

**Type:** `boolean`

**Since:** WordApi BETA

Specifies if the font is formatted as outlined. The possible values are as follows:
- true: All the text is outlined.
- false: None of the text is outlined.
- null: Returned if some, but not all, of the text is outlined.

#### Examples

**Example**: Apply outline formatting to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Apply outline formatting to the selected text
    font.outline = true;
    
    await context.sync();
});
```

---

### position

**Type:** `number`

**Since:** WordApi BETA

Specifies the position of text (in points) relative to the base line.

#### Examples

**Example**: Set the selected text to appear 6 points above the baseline as superscript positioning

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.font.position = 6;
    
    await context.sync();
});
```

---

### reflection

**Type:** `Word.ReflectionFormat`

**Since:** WordApi BETA

Returns a ReflectionFormat object that represents the reflection formatting for a shape.

#### Examples

**Example**: Add a reflection effect to the text in the first paragraph with 50% transparency and 4 points offset

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Access the reflection format
    const reflection = font.reflection;
    reflection.transparency = 0.5;
    reflection.size = 100;
    reflection.type = Word.ReflectionType.tight;
    reflection.blur = 4;
    
    await context.sync();
});
```

---

### scaling

**Type:** `number`

**Since:** WordApi BETA

Specifies the scaling percentage applied to the font.

#### Examples

**Example**: Set the font scaling to 150% for the selected text to make it appear wider

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    font.scaling = 150;
    
    await context.sync();
});
```

---

### shadow

**Type:** `boolean`

**Since:** WordApi BETA

Specifies if the font is formatted as shadowed. The possible values are as follows:
- true: All the text is shadowed.
- false: None of the text is shadowed.
- null: Returned if some, but not all, of the text is shadowed.

#### Examples

**Example**: Apply shadow formatting to the selected text in the document

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Apply shadow formatting to the selected text
    font.shadow = true;
    
    await context.sync();
});
```

---

### size

**Type:** `number`

**Since:** WordApi 1.1

Specifies a value that represents the font size in points.

#### Examples

**Example**: Change the font size of the currently selected text to 20 points.

```typescript
// Change the font size
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to change the current selection's font size.
    selection.font.size = 20;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The font size has changed.');
});
```

---

### sizeBidirectional

**Type:** `number`

**Since:** WordApi BETA

Specifies the font size in points for right-to-left text.

#### Examples

**Example**: Set the font size to 18 points for right-to-left text in the selected range

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.sizeBidirectional = 18;
    
    await context.sync();
});
```

---

### smallCaps

**Type:** `boolean`

**Since:** WordApi BETA

Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
- true: All the text has the Small Caps attribute.
- false: None of the text has the Small Caps attribute.
- null: Returned if some, but not all, of the text has the Small Caps attribute.

#### Examples

**Example**: Format selected text to display in small caps style, making lowercase letters appear as small uppercase letters

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.smallCaps = true;
    
    await context.sync();
});
```

---

### spacing

**Type:** `number`

**Since:** WordApi BETA

Specifies the spacing between characters.

#### Examples

**Example**: Set the character spacing of the selected text to 3 points to make it more spread out

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.font.spacing = 3;
    
    await context.sync();
});
```

---

### strikeThrough

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

#### Examples

**Example**: Apply strikethrough formatting to the currently selected text in the document.

```typescript
// Strike format text
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to strikethrough the font of the current selection.
    selection.font.strikeThrough = true;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The selection now has a strikethrough.');
});
```

---

### stylisticSet

**Type:** `Word.StylisticSet | "Default" | "Set01" | "Set02" | "Set03" | "Set04" | "Set05" | "Set06" | "Set07" | "Set08" | "Set09" | "Set10" | "Set11" | "Set12" | "Set13" | "Set14" | "Set15" | "Set16" | "Set17" | "Set18" | "Set19" | "Set20"`

**Since:** WordApi BETA

Specifies the stylistic set for the font.

#### Examples

**Example**: Apply the stylistic set "Set05" to the selected text's font to use alternate character forms

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.stylisticSet = Word.StylisticSet.set05;
    // Or use string literal: font.stylisticSet = "Set05";
    
    await context.sync();
});
```

---

### subscript

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

#### Examples

**Example**: Format the chemical formula "H2O" so that the "2" appears as a subscript

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert the text "H2O"
    const range = body.insertText("H2O", Word.InsertLocation.end);
    
    // Select just the "2" character (index 1)
    const subscriptRange = range.getRange().characters.items[1];
    
    // Make the "2" a subscript
    subscriptRange.font.subscript = true;
    
    await context.sync();
});
```

---

### superscript

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

#### Examples

**Example**: Format the selected text as superscript (e.g., for mathematical exponents like x²)

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.superscript = true;
    
    await context.sync();
});
```

---

### textColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA

Returns a ColorFormat object that represents the color for the font.

#### Examples

**Example**: Change the text color of the first paragraph to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    font.textColor.set("#FF0000");
    
    await context.sync();
});
```

---

### textShadow

**Type:** `Word.ShadowFormat`

**Since:** WordApi BETA

Returns a ShadowFormat object that specifies the shadow formatting for the font.

#### Examples

**Example**: Add a blue shadow with 5pt blur and 3pt offset to the selected text's font

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    const textShadow = font.textShadow;
    
    textShadow.blur = 5;
    textShadow.offsetX = 3;
    textShadow.offsetY = 3;
    textShadow.color = "blue";
    
    await context.sync();
});
```

---

### threeDimensionalFormat

**Type:** `Word.ThreeDimensionalFormat`

**Since:** WordApi BETA

Returns a ThreeDimensionalFormat object that contains 3-dimensional (3D) effect formatting properties for the font.

#### Examples

**Example**: Apply a 3D bevel effect to the selected text's font

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    // Access the 3D format properties
    const threeDFormat = font.threeDimensionalFormat;
    
    // Apply 3D bevel effects
    threeDFormat.bevelTop.type = Word.BevelType.angle;
    threeDFormat.bevelTop.width = 5;
    threeDFormat.bevelTop.height = 5;
    
    await context.sync();
});
```

---

### underline

**Type:** `Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"`

**Since:** WordApi 1.1

Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

#### Examples

**Example**: Apply single underline formatting to the currently selected text in the document.

```typescript
// Underline format text
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The selection now has an underline style.');
});
```

---

### underlineColor

**Type:** `string`

**Since:** WordApi BETA

Specifies the color of the underline for the Font object. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the underline color of selected text to red (#FF0000)

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.font.underlineColor = "#FF0000";
    
    await context.sync();
});
```

---

## Methods

### decreaseFontSize

**Kind:** `write`

Decreases the font size to the next available size.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Decrease the font size of the selected text to the next smaller available size

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.decreaseFontSize();
    
    await context.sync();
});
```

---

### increaseFontSize

**Kind:** `write`

Increases the font size to the next available size.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Increase the font size of the selected text to the next available size

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;
    
    font.increaseFontSize();
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.FontLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Font`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Font`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Font`

#### Examples

**Example**: Load and display the font name and size of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Load the font properties we want to read
    font.load("name, size");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log(`Font name: ${font.name}`);
    console.log(`Font size: ${font.size}`);
});
```

---

### reset

**Kind:** `write`

Removes manual character formatting.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove all manual character formatting from the first paragraph to restore it to the default style

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const font = firstParagraph.font;
    
    // Remove all manual character formatting
    font.reset();
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.FontUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Font` (required)

  **Returns:** `void`

#### Examples

**Example**: Format the first paragraph's text by setting multiple font properties at once (bold, size, color, and font family)

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const font = firstParagraph.font;
    
    font.set({
        bold: true,
        size: 16,
        color: "#FF0000",
        name: "Arial"
    });
    
    await context.sync();
});
```

---

### setAsTemplateDefault

**Kind:** `configure`

Sets the specified font formatting as the default for the active document and all new documents based on the active template.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set Arial font with 14pt size and blue color as the default font for the current document and all new documents based on the active template

```typescript
await Word.run(async (context) => {
    // Get the font of the body
    const bodyFont = context.document.body.font;
    
    // Set desired font properties
    bodyFont.name = "Arial";
    bodyFont.size = 14;
    bodyFont.color = "#0000FF"; // Blue
    
    // Set these font settings as the template default
    bodyFont.setAsTemplateDefault();
    
    await context.sync();
    
    console.log("Font settings have been set as template default");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Font object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FontData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.FontData`

#### Examples

**Example**: Get the font properties of the first paragraph as a plain JavaScript object for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Load font properties
    font.load("name,size,bold,italic,color");
    
    await context.sync();
    
    // Convert the Font object to a plain JavaScript object
    const fontData = font.toJSON();
    
    // Now you can use the plain object for logging, storage, or transfer
    console.log("Font properties:", fontData);
    console.log("Font name:", fontData.name);
    console.log("Font size:", fontData.size);
    console.log("Is bold:", fontData.bold);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Font`

#### Examples

**Example**: Apply bold formatting to the first paragraph's font and track it to maintain the reference across multiple sync calls for further modifications

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const font = firstParagraph.font;
    
    // Track the font object to use it across multiple sync calls
    font.track();
    
    font.bold = true;
    await context.sync();
    
    // Can safely modify the tracked font object after sync
    font.color = "blue";
    font.size = 14;
    await context.sync();
    
    // Untrack when done to free up memory
    font.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Font`

#### Examples

**Example**: Apply formatting to a paragraph's font, then untrack the font object to free memory after the changes are complete.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const font = paragraph.font;
    
    // Track the font object to work with it
    font.track();
    
    // Load and modify font properties
    font.load("name,size");
    await context.sync();
    
    font.name = "Arial";
    font.size = 14;
    font.bold = true;
    
    await context.sync();
    
    // Release the memory associated with the tracked font object
    font.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/word/word.font
