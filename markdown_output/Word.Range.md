# Word.Range

**API Set:** None None

## Description

**Package:** [word](/en-us/javascript/api/word)

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the range of the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("contentRange");
  const range: Word.Range = comment.getRange();
  range.load("text");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no range to get.");
    return;
  }

  console.log(`Comment location: ${range.text}`);
  const contentRange: Word.CommentContentRange = comment.contentRange;
  console.log("Comment content range:", contentRange);
});
```

## Properties

### bold

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Make the first paragraph in the document bold

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    range.bold = true;
    
    await context.sync();
});
```

---

### boldBidirectional

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Apply bold formatting to bidirectional text (such as Arabic or Hebrew) in the selected range

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.boldBidirectional = true;
    
    await context.sync();
});
```

---

### bookmarks

**Type:** `Word.BookmarkCollection`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get all bookmarks within the selected range and display their names in the console

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const bookmarks = range.bookmarks;
    
    bookmarks.load("items/name");
    await context.sync();
    
    console.log(`Found ${bookmarks.items.length} bookmark(s) in the selection:`);
    bookmarks.items.forEach(bookmark => {
        console.log(`- ${bookmark.name}`);
    });
});
```

---

### borders

**Type:** `Word.BorderUniversalCollection`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set all borders of the selected range to be visible with a red color and double line style

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const borders = range.borders;
    borders.load("items");
    
    await context.sync();
    
    // Set properties for all borders in the collection
    for (let i = 0; i < borders.items.length; i++) {
        borders.items[i].type = Word.BorderType.double;
        borders.items[i].color = "#FF0000";
        borders.items[i].visible = true;
    }
    
    await context.sync();
});
```

---

### case

**Type:** `Word.CharacterCase | "Next" | "Lower" | "Upper" | "TitleWord" | "TitleSentence" | "Toggle" | "HalfWidth" | "FullWidth" | "Katakana" | "Hiragana"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Change all text in the first paragraph to uppercase

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const range = paragraph.getRange();
    range.case = Word.CharacterCase.upper;
    
    await context.sync();
});
```

---

### characterWidth

**Type:** `Word.CharacterWidth | "Half" | "Full"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the character width of selected text to full-width characters (commonly used for Asian typography)

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.characterWidth = Word.CharacterWidth.full;
    
    await context.sync();
});
```

---

### combineCharacters

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the selected text has combined characters formatting and display the result in the console

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("combineCharacters");
    
    await context.sync();
    
    console.log("Combined characters enabled: " + range.combineCharacters);
});
```

---

### contentControls

**Type:** `Word.ContentControlCollection`

Gets the collection of content control objects in the range.

#### Examples

**Example**: Find all content controls in the current selection and highlight them with yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Get all content controls in the selected range
    const contentControls = range.contentControls;
    contentControls.load("items");
    
    await context.sync();
    
    // Highlight each content control with yellow background
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Range object to load and read its text property

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Access the context property from the range object
    const requestContext = range.context;
    
    // Use the context to load properties on the range
    range.load("text");
    
    await requestContext.sync();
    
    console.log("Selected text: " + range.text);
});
```

---

### disableCharacterSpaceGrid

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Disable the character space grid for the selected text range to prevent automatic character spacing adjustments

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.disableCharacterSpaceGrid = true;
    
    await context.sync();
});
```

---

### emphasisMark

**Type:** `Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Apply an emphasis mark with a solid circle above the selected text in the document

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.emphasisMark = Word.EmphasisMark.overSolidCircle;
    
    await context.sync();
});
```

---

### end

**Type:** `number`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get the end position of the first paragraph in the document and insert text at that location

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    // Load the end property
    range.load("end");
    await context.sync();
    
    // Use the end position to insert text after the paragraph
    const endPosition = range.end;
    console.log(`Paragraph ends at position: ${endPosition}`);
    
    // Insert text at the end position
    range.insertText(` [Added at position ${endPosition}]`, Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### endnotes

**Type:** `Word.NoteItemCollection`

Gets the collection of endnotes in the range.

#### Examples

**Example**: Get all endnotes in the selected range and display their reference numbers and text content in the console.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const endnotes = range.endnotes;
    
    endnotes.load("items");
    await context.sync();
    
    console.log(`Found ${endnotes.items.length} endnote(s) in the selected range:`);
    
    for (let i = 0; i < endnotes.items.length; i++) {
        const endnote = endnotes.items[i];
        endnote.load("reference, body/text");
        await context.sync();
        
        console.log(`Endnote ${i + 1}:`);
        console.log(`  Reference: ${endnote.reference}`);
        console.log(`  Text: ${endnote.body.text}`);
    }
});
```

---

### fields

**Type:** `Word.FieldCollection`

Gets the collection of field objects in the range.

#### Examples

**Example**: Get all fields in the document body and display their types and codes in the console

```typescript
await Word.run(async (context) => {
    const bodyRange = context.document.body;
    const fields = bodyRange.fields;
    
    fields.load("items");
    await context.sync();
    
    console.log(`Found ${fields.items.length} fields in the document`);
    
    for (let i = 0; i < fields.items.length; i++) {
        const field = fields.items[i];
        field.load("type, code");
        await context.sync();
        
        console.log(`Field ${i + 1}: Type = ${field.type}, Code = ${field.code}`);
    }
});
```

---

### fitTextWidth

**Type:** `number`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the fit text width to 200 points for the selected text range to compress or expand the text to fit within that width

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.fitTextWidth = 200;
    
    await context.sync();
});
```

---

### font

**Type:** `Word.Font`

Gets the text format of the range. Use this to get and set font name, size, color, and other properties.

#### Examples

**Example**: Set the font name to "Arial", size to 14, and color to blue for the first paragraph in the document.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const range = paragraph.getRange();
    
    range.font.name = "Arial";
    range.font.size = 14;
    range.font.color = "blue";
    
    await context.sync();
});
```

---

### footnotes

**Type:** `Word.NoteItemCollection`

Gets the collection of footnotes in the range.

#### Examples

**Example**: Retrieve and display the count of footnotes within the currently selected range of the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the footnotes in the selected document range.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.getSelection().footnotes;
  footnotes.load("length");
  await context.sync();

  console.log("Number of footnotes in the selected range: " + footnotes.items.length);
});
```

---

### frames

**Type:** `None`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get all frames from the selected range and log their count to the console

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const frames = range.frames;
    
    frames.load("items");
    await context.sync();
    
    console.log(`Number of frames in selection: ${frames.items.length}`);
});
```

---

### grammarChecked

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Disable grammar checking for a selected range of text in the document

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Disable grammar checking for this range
    range.grammarChecked = false;
    
    await context.sync();
});
```

---

### hasNoProofing

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the selected text range has the "no proofing" setting enabled and display the result

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("hasNoProofing");
    
    await context.sync();
    
    console.log(`No proofing enabled: ${range.hasNoProofing}`);
});
```

---

### highlightColorIndex

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Highlight the selected text in yellow to emphasize important content in the document

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.highlightColorIndex = "Yellow";
    
    await context.sync();
});
```

---

### horizontalInVertical

**Type:** `Word.HorizontalInVerticalType | "None" | "FitInLine" | "ResizeLine"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the horizontal-in-vertical text layout to fit characters within the line for the selected range

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.horizontalInVertical = Word.HorizontalInVerticalType.fitInLine;
    
    await context.sync();
});
```

---

### hyperlink

**Type:** `string`

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.

#### Examples

**Example**: Set a hyperlink on the selected text that links to "https://www.contoso.com" with a bookmark location "section2"

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Set hyperlink with address and location part
    range.hyperlink = "https://www.contoso.com#section2";
    
    await context.sync();
});
```

---

### hyperlinks

**Type:** `Word.HyperlinkCollection`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get all hyperlinks in the selected range and display their URLs in the console

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const hyperlinks = range.hyperlinks;
    
    hyperlinks.load("items");
    await context.sync();
    
    hyperlinks.items.forEach((hyperlink) => {
        console.log(hyperlink.url);
    });
});
```

---

### id

**Type:** `string`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get and display the unique identifier of the first paragraph in the document for tracking purposes

```typescript
await Word.run(async (context) => {
    // Get the first paragraph as a range
    const paragraph = context.document.body.paragraphs.getFirst();
    const range = paragraph.getRange();
    
    // Load the id property
    range.load("id");
    
    await context.sync();
    
    // Display the range ID
    console.log("Range ID: " + range.id);
});
```

---

### inlinePictures

**Type:** `Word.InlinePictureCollection`

Gets the collection of inline picture objects in the range.

#### Examples

**Example**: Get all inline pictures in the selected range and set their width to 150 pixels

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const inlinePictures = range.inlinePictures;
    
    inlinePictures.load("items");
    await context.sync();
    
    for (let i = 0; i < inlinePictures.items.length; i++) {
        inlinePictures.items[i].width = 150;
    }
    
    await context.sync();
});
```

---

### isEmpty

**Type:** `boolean`

Checks whether the range length is zero.

#### Examples

**Example**: Check if a selected range is empty and display an appropriate message to the user

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("isEmpty");
    
    await context.sync();
    
    if (range.isEmpty) {
        console.log("The selected range is empty (no content).");
    } else {
        console.log("The selected range contains content.");
    }
});
```

---

### isEndOfRowMark

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the current range is at the end of a table row and display an alert with the result

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("isEndOfRowMark");
    
    await context.sync();
    
    if (range.isEndOfRowMark) {
        console.log("The selection is at the end of a table row.");
    } else {
        console.log("The selection is not at the end of a table row.");
    }
});
```

---

### isTextVisibleOnScreen

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the selected text is currently visible on the screen and display an alert with the result

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("isTextVisibleOnScreen");
    
    await context.sync();
    
    if (range.isTextVisibleOnScreen) {
        console.log("The selected text is visible on screen");
    } else {
        console.log("The selected text is not visible on screen");
    }
});
```

---

### italic

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Make the text in the first paragraph italic

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    range.italic = true;
    
    await context.sync();
});
```

---

### italicBidirectional

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set bidirectional italic formatting on the selected text range in the document

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.italicBidirectional = true;
    
    await context.sync();
});
```

---

### kana

**Type:** `Word.Kana | "Katakana" | "Hiragana"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the kana property of selected text to Hiragana to specify how Japanese phonetic characters should be displayed

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.kana = Word.Kana.hiragana;
    
    await context.sync();
});
```

---

### languageDetected

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the language of the selected text range has been automatically detected and display the result

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("languageDetected");
    
    await context.sync();
    
    if (range.languageDetected) {
        console.log("Language has been automatically detected for this range");
    } else {
        console.log("Language has not been automatically detected for this range");
    }
});
```

---

### languageId

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the language of the selected text to French

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.languageId = "French";
    
    await context.sync();
});
```

---

### languageIdFarEast

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the Far East language formatting to Japanese for the selected text range in a Word document

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Set the Far East language to Japanese
    range.languageIdFarEast = "Japanese";
    
    // Sync to apply the changes
    await context.sync();
    
    console.log("Far East language set to Japanese for the selected range");
});
```

---

### languageIdOther

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set the language ID for non-Latin text in the selected range to Japanese

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.languageIdOther = "Japanese";
    
    await context.sync();
    console.log("Language ID for non-Latin text set to Japanese");
});
```

---

### listFormat

**Type:** `Word.ListFormat`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get the list formatting information from the selected text range and display whether it's part of a numbered or bulleted list

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const listFormat = range.listFormat;
    
    listFormat.load("listLevelNumber, listType");
    await context.sync();
    
    console.log("List Level: " + listFormat.listLevelNumber);
    console.log("List Type: " + listFormat.listType);
});
```

---

### lists

**Type:** `Word.ListCollection`

Gets the collection of list objects in the range.

#### Examples

**Example**: Get all lists in the document body and log the count of lists found

```typescript
await Word.run(async (context) => {
    const bodyRange = context.document.body;
    const lists = bodyRange.lists;
    
    lists.load("items");
    await context.sync();
    
    console.log(`Number of lists in the document: ${lists.items.length}`);
});
```

---

### pages

**Type:** `Word.PageCollection`

Gets the collection of pages in the range.

#### Examples

**Example**: Get the number of pages in the current document's body range and display it in the console.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const pages = body.getRange().pages;
    
    pages.load("items");
    await context.sync();
    
    console.log(`The document contains ${pages.items.length} page(s)`);
});
```

---

### paragraphs

**Type:** `Word.ParagraphCollection`

Gets the collection of paragraph objects in the range.

#### Examples

**Example**: Get all paragraphs in the selected range and highlight each paragraph with yellow color

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const paragraphs = range.paragraphs;
    
    paragraphs.load("items");
    await context.sync();
    
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### parentBody

**Type:** `Word.Body`

Gets the parent body of the range.

#### Examples

**Example**: Get the text content of the entire body that contains a selected range of text

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Get the parent body of the range
    const parentBody = range.parentBody;
    
    // Load the text property of the parent body
    parentBody.load("text");
    
    await context.sync();
    
    // Display the parent body's text
    console.log("Parent body text: " + parentBody.text);
});
```

---

### parentContentControl

**Type:** `Word.ContentControl`

Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.

#### Examples

**Example**: Toggle the checked state of a checkbox content control in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Toggles the isChecked property of the first checkbox content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.checkBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,checkboxContentControl/isChecked");

  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,checkboxContentControl/isChecked");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
      console.warn("No checkbox content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  const isCheckedBefore = selectedContentControl.checkboxContentControl.isChecked;
  console.log("isChecked state before:", `id: ${selectedContentControl.id} ... isChecked: ${isCheckedBefore}`);
  selectedContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
  selectedContentControl.load("id,checkboxContentControl/isChecked");
  await context.sync();

  console.log(
    "isChecked state after:",
    `id: ${selectedContentControl.id} ... isChecked: ${selectedContentControl.checkboxContentControl.isChecked}`
  );
});
```

---

### parentContentControlOrNullObject

**Type:** `Word.ContentControl`

Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a selected range is inside a content control and highlight the content control's title if it exists.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const parentContentControl = range.parentContentControlOrNullObject;
    
    parentContentControl.load("isNullObject, title, tag");
    await context.sync();
    
    if (!parentContentControl.isNullObject) {
        console.log(`Range is inside content control: "${parentContentControl.title}"`);
        console.log(`Content control tag: "${parentContentControl.tag}"`);
        
        // Highlight the content control
        parentContentControl.appearance = Word.ContentControlAppearance.tags;
    } else {
        console.log("Range is not inside a content control");
    }
    
    await context.sync();
});
```

---

### parentTable

**Type:** `Word.Table`

Gets the table that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table.

#### Examples

**Example**: Get the table containing a selected range and highlight all cells in that table with yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Get the parent table that contains this range
    const parentTable = range.parentTable;
    
    // Load the table's cells
    parentTable.load("cells");
    
    await context.sync();
    
    // Highlight all cells in the parent table
    for (let i = 0; i < parentTable.cells.items.length; i++) {
        parentTable.cells.items[i].body.font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### parentTableCell

**Type:** `Word.TableCell`

Gets the table cell that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table cell.

#### Examples

**Example**: Get the table cell containing a selected range and highlight it with a yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Get the parent table cell that contains this range
    const tableCell = range.parentTableCell;
    
    // Load the cell's shading property
    tableCell.load("shading");
    
    await context.sync();
    
    // Highlight the cell with yellow background
    tableCell.shading.backgroundPatternColor = "yellow";
    
    await context.sync();
});
```

---

### parentTableCellOrNullObject

**Type:** `Word.TableCell`

Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if the selected range is inside a table cell and highlight the cell yellow if it is, otherwise show a message that the selection is not in a table.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const tableCell = range.parentTableCellOrNullObject;
    
    tableCell.load("isNullObject");
    await context.sync();
    
    if (tableCell.isNullObject) {
        console.log("The selection is not inside a table cell.");
    } else {
        tableCell.shadingColor = "yellow";
        console.log("The table cell containing the selection has been highlighted.");
    }
    
    await context.sync();
});
```

---

### parentTableOrNullObject

**Type:** `Word.Table`

Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if the selected range is inside a table, and if so, highlight the entire parent table with a yellow background color.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const parentTable = range.parentTableOrNullObject;
    
    // Load the isNullObject property to check if range is in a table
    parentTable.load("isNullObject");
    await context.sync();
    
    if (!parentTable.isNullObject) {
        // Range is inside a table, so highlight it
        parentTable.shadingColor = "#FFFF00"; // Yellow background
        await context.sync();
        console.log("Parent table highlighted");
    } else {
        console.log("Selected range is not inside a table");
    }
});
```

---

### sections

**Type:** `Word.SectionCollection`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get all sections within a selected range and display the count of sections found

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const sections = range.sections;
    
    sections.load("items");
    await context.sync();
    
    console.log(`Number of sections in the selected range: ${sections.items.length}`);
});
```

---

### shading

**Type:** `Word.ShadingUniversal`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Apply yellow background shading to the selected text range in the document

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.shading.backgroundPatternColor = "yellow";
    
    await context.sync();
});
```

---

### shapes

**Type:** `Word.ShapeCollection`

Gets the collection of shape objects anchored in the range, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

#### Examples

**Example**: Get all shapes anchored in the selected range and log their count and types to the console.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    const shapes = range.shapes;
    
    shapes.load("items/type");
    await context.sync();
    
    console.log(`Found ${shapes.items.length} shape(s) in the selected range`);
    shapes.items.forEach((shape, index) => {
        console.log(`Shape ${index + 1}: ${shape.type}`);
    });
});
```

---

### showAll

**Type:** `None`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Show all formatting marks (such as spaces, tabs, and paragraph marks) in the selected text range to make document structure visible

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.showAll = true;
    
    await context.sync();
});
```

---

### spellingChecked

**Type:** `boolean`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Mark a range of text as already spell-checked to prevent Word from displaying spelling errors for that content

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.spellingChecked = true;
    
    await context.sync();
    console.log("The selected text has been marked as spell-checked.");
});
```

---

### start

**Type:** `number`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get the starting character position of the first paragraph in the document and display it in the console

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    range.load("start");
    
    await context.sync();
    
    console.log(`The paragraph starts at character position: ${range.start}`);
});
```

---

### storyLength

**Type:** `number`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get the total number of characters in the story that contains the selected range and display it to the user

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("storyLength");
    
    await context.sync();
    
    console.log(`The story contains ${range.storyLength} characters.`);
});
```

---

### storyType

**Type:** `Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Check if the selected text is in the main document body or in a header/footer, and display the story type to the user.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("storyType");
    
    await context.sync();
    
    console.log(`Selected text is in: ${range.storyType}`);
    
    if (range.storyType === "MainText") {
        console.log("This is main document content");
    } else if (range.storyType === "PrimaryHeader" || range.storyType === "FirstPageHeader") {
        console.log("This is header content");
    } else if (range.storyType === "PrimaryFooter" || range.storyType === "FirstPageFooter") {
        console.log("This is footer content");
    }
});
```

---

### style

**Type:** `string`

Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

#### Examples

**Example**: Apply a custom style named "CustomHeading" to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    range.style = "CustomHeading";
    
    await context.sync();
});
```

---

### styleBuiltIn

**Type:** `Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"`

Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

#### Examples

**Example**: Insert a heading with the text "This is a sample Heading 1 Title!!" at the beginning of the document body and apply the built-in Heading 1 style to it.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

await Word.run(async (context) => {
    const header: Word.Range = context.document.body.insertText("This is a sample Heading 1 Title!!\n",
        "Start" /*this means at the beginning of the body */);
    header.styleBuiltIn = Word.BuiltInStyleName.heading1;

    await context.sync();
});
```

---

### tableColumns

**Type:** `Word.TableColumnCollection`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Get the count of table columns that intersect with the selected range and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Get the table columns that intersect with this range
    const tableColumns = range.tableColumns;
    tableColumns.load("count");
    
    await context.sync();
    
    console.log(`Number of table columns in range: ${tableColumns.count}`);
});
```

---

### tables

**Type:** `Word.TableCollection`

Gets the collection of table objects in the range.

#### Examples

**Example**: Get all tables in the current selection and highlight the first table by setting its shading color to light yellow.

```typescript
await Word.run(async (context) => {
    // Get the current selection range
    const range = context.document.getSelection();
    
    // Get all tables in the selected range
    const tables = range.tables;
    tables.load("items");
    
    await context.sync();
    
    // Check if there are any tables in the range
    if (tables.items.length > 0) {
        // Highlight the first table with light yellow shading
        const firstTable = tables.items[0];
        firstTable.shadingColor = "#FFFFE0";
        
        await context.sync();
        console.log(`Found ${tables.items.length} table(s) in the selection.`);
    } else {
        console.log("No tables found in the selected range.");
    }
});
```

---

### text

**Type:** `string`

Gets the text of the range.

#### Examples

**Example**: Read and display the text content from the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    range.load("text");
    await context.sync();
    
    console.log("Paragraph text: " + range.text);
});
```

---

### twoLinesInOne

**Type:** `Word.TwoLinesInOneType | "None" | "NoBrackets" | "Parentheses" | "SquareBrackets" | "AngleBrackets" | "CurlyBrackets"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Format selected text to display as two lines in one with parentheses brackets

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.twoLinesInOne = "Parentheses";
    
    await context.sync();
});
```

---

### underline

**Type:** `Word.Underline | "None" | "Single" | "Words" | "Double" | "Dotted" | "Thick" | "Dash" | "DotDash" | "DotDotDash" | "Wavy" | "WavyHeavy" | "DottedHeavy" | "DashHeavy" | "DotDashHeavy" | "DotDotDashHeavy" | "DashLong" | "DashLongHeavy" | "WavyDouble"`

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Apply a wavy underline style to the selected text in the document

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.underline = Word.Underline.wavy;
    
    await context.sync();
});
```

---

## Methods

### clear

**Kind:** `delete`

Clears the contents of the range object. The user can perform the undo operation on the cleared content.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Clear the contents of the currently selected text range in the Word document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Cleared the selection (range object)');
});
```

---

### compareLocationWith

Compares this range's location with another range's location.

#### Signature

**Parameters:**
- `range`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Determine the spatial relationship between the first paragraph and the second paragraph in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/compare-location.yaml

// Compares the location of one paragraph in relation to another paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("items");

  await context.sync();

  const firstParagraphAsRange: Word.Range = paragraphs.items[0].getRange();
  const secondParagraphAsRange: Word.Range = paragraphs.items[1].getRange();

  const comparedLocation = firstParagraphAsRange.compareLocationWith(secondParagraphAsRange);

  await context.sync();

  const locationValue: Word.LocationRelation = comparedLocation.value;
  console.log(`Location of the first paragraph in relation to the second paragraph: ${locationValue}`);
});
```

---

### delete

**Kind:** `delete`

Deletes the range and its content from the document.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete the currently selected text or content from the Word document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Deleted the selection (range object)');
});
```

---

### detectLanguage

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Detect the language of the selected text in the document and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Detect the language of the selected text
    const languageInfo = range.detectLanguage();
    
    // Load the language ID property
    languageInfo.load("id");
    
    await context.sync();
    
    // Display the detected language ID
    console.log(`Detected language ID: ${languageInfo.id}`);
});
```

---

### expandTo

Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges don't have a union.

#### Signature

**Parameters:**
- `range`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display all complete sentences from the current insertion point to the end of the paragraph.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // Get the complete sentence (as range) associated with the insertion point.
  const sentences: Word.RangeCollection = context.document
    .getSelection()
    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
  sentences.load("$none");
  await context.sync();

  // Expand the range to the end of the paragraph to get all the complete sentences.
  const sentencesToTheEndOfParagraph: Word.RangeCollection = sentences.items[0]
    .getRange()
    .expandTo(
      context.document
        .getSelection()
        .paragraphs.getFirst()
        .getRange(Word.RangeLocation.end)
    )
    .getTextRanges(["."], false /* Don't trim spaces*/);
  sentencesToTheEndOfParagraph.load("text");
  await context.sync();

  for (let i = 0; i < sentencesToTheEndOfParagraph.items.length; i++) {
    console.log(sentencesToTheEndOfParagraph.items[i].text);
  }
});
```

---

### expandToOrNullObject

Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. If the two ranges don't have a union, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Signature

**Parameters:**
- `range`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Expand a selected range to include both the selected text and the first paragraph in the document, or detect if they cannot be combined into a single range.

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const firstParagraph = context.document.body.paragraphs.getFirst().getRange();
    
    // Expand the selection to cover both the selected text and the first paragraph
    const expandedRange = selection.expandToOrNullObject(firstParagraph);
    
    expandedRange.load("isNullObject, text");
    await context.sync();
    
    if (expandedRange.isNullObject) {
        console.log("The ranges cannot be combined (no union exists)");
    } else {
        // Highlight the expanded range
        expandedRange.font.highlightColor = "yellow";
        console.log("Expanded range text:", expandedRange.text);
    }
    
    await context.sync();
});
```

---

### getBookmarks

**Kind:** `read`

Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.

#### Signature

**Parameters:**
- `includeHidden`: `None` (required)
- `includeAdjacent`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get all visible bookmarks in the selected range and display them in the console

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    // Get only visible bookmarks (not hidden ones starting with underscore)
    // and exclude adjacent bookmarks
    const bookmarks = range.getBookmarks(false, false);
    
    context.load(bookmarks);
    await context.sync();
    
    console.log("Visible bookmarks in selection:");
    bookmarks.value.forEach(bookmark => {
        console.log(bookmark);
    });
});
```

---

### getComments

**Kind:** `read`

Gets comments associated with the range.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display all comments associated with the currently selected content in the Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the comments in the selected content.
await Word.run(async (context) => {
  const comments: Word.CommentCollection = context.document.getSelection().getComments();

  // Load objects to log in the console.
  comments.load();
  await context.sync();

  console.log("Comments:", comments);
});
```

---

### getContentControls

**Kind:** `read`

Gets the currently supported content controls in the range.

#### Signature

**Parameters:**
- `options`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Delete the first checkbox content control found in the current selection or its parent container, warning if no checkbox is selected.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Deletes the first checkbox content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.checkBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id");

  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
      console.warn("No checkbox content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  console.log(`About to delete checkbox content control with id: ${selectedContentControl.id}`);
  selectedContentControl.delete(false);
  await context.sync();

  console.log("Deleted checkbox content control.");
});
```

---

### getHtml

**Kind:** `read`

Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve the HTML representation of the currently selected text in the Word document and output it to the console.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to get the HTML of the current selection.
    const html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The HTML read from the document was: ' + html.value);
});
```

---

### getHyperlinkRanges

**Kind:** `read`

Gets hyperlink child ranges within the range.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and log all hyperlinks from the entire document body to the console.

```typescript
await Word.run(async (context) => {
    // Get the entire document body.
    const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);

    // Get all the ranges that only consist of hyperlinks.
    const hyperLinks = bodyRange.getHyperlinkRanges();
    hyperLinks.load("hyperlink");
    await context.sync();

    // Log each hyperlink.
    hyperLinks.items.forEach((linkRange) => {
        console.log(linkRange.hyperlink);
    });
});
```

---

### getNextTextRange

**Kind:** `read`

Gets the next text range by using punctuation marks and/or other ending marks. Throws an `ItemNotFound` error if this text range is the last one.

#### Signature

**Parameters:**
- `endingMarks`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find and highlight the first sentence in the document by getting the text range from the start of the document to the first period.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's range
    const body = context.document.body;
    const firstParagraph = body.paragraphs.getFirst();
    const paragraphRange = firstParagraph.getRange();
    
    // Get the text range up to the first period (sentence ending)
    const firstSentence = paragraphRange.getNextTextRange(["."], true);
    
    // Highlight the first sentence
    firstSentence.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### getNextTextRangeOrNullObject

**Kind:** `read`

Gets the next text range by using punctuation marks and/or other ending marks. If this text range is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Signature

**Parameters:**
- `endingMarks`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Split a paragraph into sentences by finding text ranges separated by periods, and highlight every other sentence in yellow.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("text");
    await context.sync();
    
    // Get the first text range (from start to first period)
    let currentRange = paragraph.getRange().getNextTextRangeOrNullObject(["."], true);
    currentRange.load("text, isNullObject");
    await context.sync();
    
    let count = 0;
    // Loop through all sentences
    while (!currentRange.isNullObject) {
        // Highlight every other sentence
        if (count % 2 === 0) {
            currentRange.font.highlightColor = "yellow";
        }
        
        // Get the next sentence
        currentRange = currentRange.getNextTextRangeOrNullObject(["."], true);
        currentRange.load("text, isNullObject");
        await context.sync();
        count++;
    }
});
```

---

### getOoxml

**Kind:** `read`

Gets the OOXML representation of the range object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve the OOXML representation of the currently selected text in the Word document and log it to the console.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to get the OOXML of the current selection.
    const ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The OOXML read from the document was:  ' + ooxml.value);
});
```

---

### getRange

**Kind:** `read`

Clones the range, or gets the starting or ending point of the range as a new range.

#### Signature

**Parameters:**
- `rangeLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a dropdown list content control at the end of the current selection in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Places a dropdown list content control at the end of the selection.
await Word.run(async (context) => {
  let selection = context.document.getSelection();
  selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.dropDownList);
  await context.sync();

  console.log("Dropdown list content control inserted at the end of the selection.");
});
```

---

### getReviewedText

**Kind:** `read`

Gets reviewed text based on ChangeTrackingVersion selection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Retrieve and compare the original and current versions of the selected text to show changes made during document review.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the reviewed text.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection();
  const before = range.getReviewedText(Word.ChangeTrackingVersion.original);
  const after = range.getReviewedText(Word.ChangeTrackingVersion.current);

  await context.sync();

  console.log("Reviewed text (before):", before.value, "Reviewed text (after):", after.value);
});
```

---

### getTextRanges

**Kind:** `read`

Gets the text child ranges in the range by using punctuation marks and/or other ending marks.

#### Signature

**Parameters:**
- `endingMarks`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display all complete sentences from the current insertion point to the end of the paragraph.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // Get the complete sentence (as range) associated with the insertion point.
  const sentences: Word.RangeCollection = context.document
    .getSelection()
    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
  sentences.load("$none");
  await context.sync();

  // Expand the range to the end of the paragraph to get all the complete sentences.
  const sentencesToTheEndOfParagraph: Word.RangeCollection = sentences.items[0]
    .getRange()
    .expandTo(
      context.document
        .getSelection()
        .paragraphs.getFirst()
        .getRange(Word.RangeLocation.end)
    )
    .getTextRanges(["."], false /* Don't trim spaces*/);
  sentencesToTheEndOfParagraph.load("text");
  await context.sync();

  for (let i = 0; i < sentencesToTheEndOfParagraph.items.length; i++) {
    console.log(sentencesToTheEndOfParagraph.items[i].text);
  }
});
```

---

### getTrackedChanges

**Kind:** `read`

Gets the collection of the TrackedChange objects in the range.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get all tracked changes in the current document selection and display their count and types in the console.

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Get tracked changes in the selected range
    const trackedChanges = range.getTrackedChanges();
    
    // Load properties of the tracked changes
    trackedChanges.load("items");
    
    await context.sync();
    
    // Display information about tracked changes
    console.log(`Found ${trackedChanges.items.length} tracked change(s) in selection`);
    
    trackedChanges.items.forEach((change, index) => {
        change.load("type");
    });
    
    await context.sync();
    
    trackedChanges.items.forEach((change, index) => {
        console.log(`Change ${index + 1}: ${change.type}`);
    });
});
```

---

### highlight

Highlights the range temporarily without changing document content. To highlight the text permanently, set the range's Font.HighlightColor.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Temporarily highlight all occurrences of the word "important" in the document to draw attention during review

```typescript
await Word.run(async (context) => {
    // Search for all occurrences of "important"
    const searchResults = context.document.body.search("important", { matchCase: false });
    searchResults.load("items");
    
    await context.sync();
    
    // Temporarily highlight each occurrence
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].highlight();
    }
    
    await context.sync();
});
```

---

### insertBookmark

**Kind:** `create`

Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.

#### Signature

**Parameters:**
- `name`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a bookmark named "ImportantSection" at the currently selected text range in the document

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Insert a bookmark on the selected range
    range.insertBookmark("ImportantSection");
    
    await context.sync();
    
    console.log("Bookmark 'ImportantSection' has been inserted at the selected range.");
});
```

---

### insertBreak

**Kind:** `create`

Inserts a break at the specified location in the main document.

#### Signature

**Parameters:**
- `breakType`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a page break immediately after the currently selected text in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert a page break after the selected text.
    range.insertBreak(Word.BreakType.page, Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Inserted a page break after the selected text.');
});
```

---

### insertCanvas

**Kind:** `create`

Inserts a floating canvas in front of text with its anchor at the beginning of the range.

#### Signature

**Parameters:**
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a floating canvas with a rectangle shape at the beginning of the first paragraph in the document.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    // Define options for the canvas shape
    const insertShapeOptions: Word.InsertShapeOptions = {
        shapeType: Word.ShapeType.rectangle,
        width: 200,
        height: 100
    };
    
    // Insert a floating canvas at the beginning of the range
    range.insertCanvas(insertShapeOptions);
    
    await context.sync();
});
```

---

### insertComment

**Kind:** `create`

Insert a comment on the range.

#### Signature

**Parameters:**
- `commentText`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Add a comment with user-provided text to the currently selected content in the Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Sets a comment on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("comment-text") as HTMLInputElement).value;
  const comment: Word.Comment = context.document.getSelection().insertComment(text);

  // Load object to log in the console.
  comment.load();
  await context.sync();

  console.log("Comment inserted:", comment);
});
```

---

### insertContentControl

**Kind:** `create`

Wraps the Range object with a content control.

#### Signature

**Parameters:**
- `contentControlType`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find all instances of "Contractor" in the document, make them bold, wrap each in a content control with tag "customer" and a numbered title, but only if these content controls don't already exist.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

// Simulates creation of a template. First searches the document for instances of the string "Contractor",
// then changes the format  of each search result,
// then wraps each search result within a content control,
// finally sets a tag and title property on each content control.
await Word.run(async (context) => {
    const results: Word.RangeCollection = context.document.body.search("Contractor");
    results.load("font/bold");

    // Check to make sure these content controls haven't been added yet.
    const customerContentControls: Word.ContentControlCollection = context.document.contentControls.getByTag("customer");
    customerContentControls.load("text");
    await context.sync();

  if (customerContentControls.items.length === 0) {
    for (let i = 0; i < results.items.length; i++) { 
        results.items[i].font.bold = true;
        let cc: Word.ContentControl = results.items[i].insertContentControl();
        cc.tag = "customer";  // This value is used in the next step of this sample.
        cc.title = "Customer Name " + i;
    }
  }
    await context.sync();
});
```

---

### insertEndnote

**Kind:** `create`

Inserts an endnote. The endnote reference is placed after the range.

#### Signature

**Parameters:**
- `insertText`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert an endnote at the end of the first paragraph with reference text "See additional notes in appendix"

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    // Insert an endnote with the specified text
    range.insertEndnote("See additional notes in appendix");
    
    await context.sync();
});
```

---

### insertField

**Kind:** `create`

Inserts a field at the specified location.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `insertLocation`: `None` (required)
  - `fieldType`: `None` (required)
  - `text`: `None` (required)
  - `removeFormatting`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `insertLocation`: `None` (required)
  - `fieldType`: `None` (required)
  - `text`: `None` (required)
  - `removeFormatting`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Insert a date field with custom formatting before the current selection and display its code and result in the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Inserts a Date field before selection.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection().getRange();

  const field: Word.Field = range.insertField(Word.InsertLocation.before, Word.FieldType.date, '\\@ "M/d/yyyy h:mm am/pm"', true);

  field.load("result,code");
  await context.sync();

  if (field.isNullObject) {
    console.log("There are no fields in this document.");
  } else {
    console.log("Code of the field: " + field.code, "Result of the field: " + JSON.stringify(field.result));
  }
});
```

---

### insertFileFromBase64

**Kind:** `create`

Inserts a document at the specified location.

#### Signature

**Parameters:**
- `base64File`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a base64-encoded .docx file at the beginning of the currently selected range in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Added base64 encoded text to the beginning of the range.');
});
```

---

### insertFootnote

**Kind:** `create`

Inserts a footnote. The footnote reference is placed after the range.

#### Signature

**Parameters:**
- `insertText`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a footnote with specified text at the currently selected content in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Sets a footnote on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("input-footnote") as HTMLInputElement).value;
  const footnote: Word.NoteItem = context.document.getSelection().insertFootnote(text);
  await context.sync();

  console.log("Inserted footnote.");
});
```

---

### insertGeometricShape

**Kind:** `create`

Inserts a geometric shape in front of text with its anchor at the beginning of the range.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `geometricShapeType`: `None` (required)
  - `insertShapeOptions`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `geometricShapeType`: `None` (required)
  - `insertShapeOptions`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Insert a blue rectangle shape at the beginning of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    const shape = range.insertGeometricShape(
        Word.GeometricShapeType.rectangle,
        {
            width: 100,
            height: 50,
            left: 0,
            top: 0
        }
    );
    
    shape.fill.setSolidColor("blue");
    
    await context.sync();
});
```

---

### insertHtml

**Kind:** `create`

Inserts HTML at the specified location.

#### Signature

**Parameters:**
- `html`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert bold HTML text at the beginning of the currently selected range in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('HTML added to the beginning of the range.');
});
```

---

### insertInlinePictureFromBase64

**Kind:** `create`

Inserts a picture at the specified location.

#### Signature

**Parameters:**
- `base64EncodedImage`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert an inline picture from Base64-encoded image data at the start of the first text box in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a picture at the start of the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.load("type/body");
  await context.sync();

  const startRange: Word.Range = firstShapeWithTextBox.body.getRange(Word.RangeLocation.start);
  const newPic: Word.InlinePicture = startRange.insertInlinePictureFromBase64(
    getPictureBase64(),
    Word.InsertLocation.start
  );
  newPic.load();
  await context.sync();

  console.log("New inline picture properties:", newPic);
});

...

// Returns Base64-encoded image data for a sample picture.
const pictureBase64 =
"iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAMAAAAJbSJIAAABblBMVEX+7tEYMFlyg5v8zHXVgof///+hrL77qRnIWmBEWXq6MDgAF0/i1b//8dP+79QKJ1MAIFL8yWpugZz/+O/VzLwzTXR+jaP/z3PHzdjNaWvuxrLFT1n8znmMj5fFTFP25OHlsa2wqqJGW3z7pgCbqsH936oAJlWnssRzdoLTd1HTfINbY3a7tar90IxJVG0AH1ecmJH//90gN14AFU/nxInHVFL80YQAD03qv3LUrm7cwJLWjoLenpPRdXTQgoj15sz+57/7szr93KPbiWjUvZj95LnwzLmMX3L8wmz7rib8xnP8vVz91JT8ukvTz8i8vsORkJKvsLIAD1YwPViWnKZVYHbKuqHjwo3ur2/Pa2O+OTvHVETfj1tybm9qdYlsYlnkmmC0DSPirpvAq4bj5uuono7tu5vgpannnX3ksbSKg5bv0tTclJNFSlyZgpPqwsW4go2giWdbWV+3mmuWgpRcbolURmReS2embHkiRHBcZ6c8AAALcElFTVR4nO3di1cTVx4H8AyThmC484ghFzSxEDRhIRBIMEFQA1qoVhAqYBVd3UXcri1dd7fLdv3vdybJZF73zr2TufPyzPccew49hc6H331nZkylkiRJkiRJkiRJkiRJkiRJkiRJkiRJkiQJ6wj2hH1JLKNo9p/sPB3X8rRUau/f2f56kML2k/n5+XFDSjzPQ7l95+swCqkfzDy1hnwvsLT9FRCF1I7Fpwt5Xt6PfRmF1LgNaBAqZdyNOVGwV9AkVMq4HOshR3iCAJqFalONr1HYRQGtQsXYvrONmjKj7xae0QnVuaO0/OiOlv3lfqI/1G4jgShhnzkIfzA/SNgAUoR9d0I9g/9wfjtsAiHocWZ8fIckLA1ad/SFB0jg+AGxhgNi9FvpU7TwGVHIl+QdtR9GfaTBCOdlIlA18vIzPqZC8kCjZT+mQnI31HInpkKqRqpGDhtADFpInCuGaUe9hBghrY+Xo7+xQgnn6Xth9EuIFNIPpDDsy6cISvg1tVGkkB4Y+ZlCjU34lBrIx6GCitAyyOzQ8mA7+nvfXixCigV33xf9tYwWg3B+/ICnAsbrKFwY8nae0figwnsUq3M34aCXZ3KphPa12+2SWjYZ8v0Pa1Jx4ikRSv1ga2Y8MIzH6aElAqFlRn/vQApRuB32FXoNSRiTad0hgkxI5E8piLlOStgX6DnfkBL7GhKFsS8iUfhN2FfoNWRh3ItIFsa9iBTCmBeRQhjz4ZRGGG8ilfB6jInEVVs/MTj5xUWwbSbUQNs2sZ2Kq9EilNup60qj3LUReT4mR2u2mIXyrtbx2nbjI/P+HpgTFoAYAQlU0rYJYXt3aASg+/zw8HBlkKWFuW5UkSbhsnH4RHxIKmtG8Lx2O5PJ1DhxkKqUW+hGk2gUyoJxhniE6Ivq3W0pAXQPVZ8ibHJ6qrl6JImmGppnecwn3XK7kBnEJOS4zlEUiUZh2zzLI4UQrv94GyPkOnMRJBqFyzghHKa0qfvsQk6KYF90bqUb93pZ72fz5Y+3DT6EsFqOtlC+bh1pXjSUtCq3tWTMsQm5VrSF/L6lkW7k1KsWM7jUjq3CXCFyRPOMb9hpLCtfb7TUvlWsYYUrVqG0Gm2hgbjfG2c61erxCRaYqS2J1o4YvQnDuvJeFtSV9zbfm+7hSTGD9ykpVq3ChagL1d1T/09PWLeOLdZYW2kchKbpfZMgrJ2K8RbyPKGEmRMp5kL40mURYyckFzHTjLkQrpPGmhMx3kIe/kRqp0Ux3kKlihlnY+2EE6MuhIYgiPxL25LbTMysSFEWQvjq8evs3Wu9nL15+4MdCdsvM47IWvG42q9j9c+RE4JXr29ms5pQzVtkHX9S94aG2JrquxVRqlZz7yN2Og5SW6rPJLz2BtkdlbTXN797qeS7zXX7YqdWq2VOTk7monTzBgDgPNsHmoTX3qBO2TRmP9hJpA7lRyESzafUe/c1n0V47S/EARa3YL1dh2He/Q26W2ruq9l6kL059FmFZ7giDoW41Zwq5PmwgClw/lf1+hWaEYcQXntFEMrPpzEpqBuv0EabvjCLikX4liA0n6zazpFhWLdIK8KzW0hgNmsW/sm5mcrbzsLQnjQBXWvj1HPmRshjgdpnAaFNGVhg9pYLofFDOIxQDunzVHAfX0QXwhIeOPw8J6TBBnRx3dAy1jgKzUfjGGEUi3hGKZSBA1D/TC6sngjSVEQHIfxQdMqq9p2hPbgHtvAN9YxCCD/mxwzJ54tF5R/617owtOUpuDGDLeMZSQhLRybg2LTaMi/G8nYhXwpvdQpupO3LtsFwc+YkhHBzzAzUel8RIQzzOQYAUnvnWw9mZlTUayvy7q2zM5QQ8ptlsy9/oQkv8nZhyE+3DW/zAfAtopaPrUJlR/jRUr+xsaI+hBYRwohshQX4mCyEGx+KeatvLF/ThYd5uzC8jmiKAO/esscoVMq3auepmkNdOI0QRuSRKaH0LSJd/TrhehnpUzQZXVhDCGFEHijadVyZwPUjjE/l6N+AGEvD2yVaglxkDoRww8FnLGINNZaGN+ebIqCAg506/9HJZ+iJ06gZPyqDKRLYE9qmdxSxOH1xMV1ErdqULEdAiNsmCDLkV4m+HilvqrNJGIHjbzD76dMsKn+D6+QCIsGREgJwf1HPw59/1r/4+4eRfBETgu7lYlrL4rdq4/yk/YtfRgSahaEuagDozuq+AVAjPhyRFyEhAHuzi0bgJ22IWfQGtAoBMv7zurNpo08R/qoJL70BLUJQL6Pi72226kdOZp5F6AloERZazQlbpqqnPgoV36XNZ26lnoAWIcdxUxWrsMk1/LuBUfXZeL0MgJ8Xf2Eo/E20EyvqHUadgj+9EqTuY3zp9GUP+OuDf4w6TdiF8H3/Dg0TsTK4hao+TIGdEewh2qehoX7+fLn4T49A42nivxqDO1AmKjYgJw2TqzJ6EMWpgH2i4vc2ypiE8J4GNBArtjvfuX6bZQF0LKAWj53QKNxoGAwTlUpF+TOBBHLiCgMhuEHhS3tuowbhsemGvuaUOk0gfeptRl3vQEILZVZCTQj/bb0B3CmSZyElkEEJB0J9lKHKsddWCnCTIPsS9oXw95YboOe7/SgrmH7IoIR94T1XFeQ6k96EYJYOmPY62Q+FJVc+ruPxMRtlmqADMmmkPeFv1gdpHJuo5PmZRUpfOs2ihKrwvUR2aRE7np8epu2EbEZSVfh7jt7XWimseQVSt1FGwrF3tBNhVWotMVh1g0vqRvofJsA8uQ9WG51WQ1wp11k8we+ihGwGmjH0ytPYMnPlgrqEYbQxpO+FaY97+0GwS88h8HiS7UkUPZCJcILYRptsT6HcNFIWwisisMX4MWHq5QwbIRnI/HkTFyMpCyHJx2QjaBG6KKH3AwziMMrlmL9UohukcIrYRpmcVpjiaqDxKqyQp3rWw0ywQvIo48djbQEKKRZrnMTa51boZeGdJ48yXMOHd9eMKLyqTDVFlyEDOebDzIjCqymqy3UfyY+XSNEdAxuFFc4fnpIOe59bIdWAP3o8n4l6F141/QSKvjwB7Ur4vZ8+LgI1/K/PQC4XstB3INfw4wVS9EL/gf50RGrhH/4DlWbq8dMJL0K/B5l+/HifBKXwf4EAlTmf9QafWkixamYSH17lRicMpo1yfmzxKYVBAZWxhnkzpRIGVkI/3qlIJQzMp3RE5ntgGmFQA6ka9u9UpBH+ERzQh9e3gm52BpMh3c2NPZ6FPhy2YZ9pzmYfBN5IfRGe4x9Nz84EPJL69B4whyL2iEF2Q39Wpnv4h+97RNt7gOMmVIZTh3aaDW5N2k9zjb1QqSL+/QLZmYeBApVlmy9HGeD8wU1MsotBDjT+vShafb/ADXT2XNygxSKiL8A+Ep1uwMLqgh890SlBC7ncasDErqt7eVmkVQ70L2sBddc11J8EaeRGWtNKTfVvpAnqmT3gfsJfG6ZbKEujGTunC6tz1tQ93g2G/qUtub/CJS0LR3WQKo/WysWqZE/reG5Uo4qZLNh+aXNlcYQS6B/7VhvS0Vqd/nZZchrHIx0aK7q5dxNThoiDX5r3raF0nKqzHKtEyf1JDgD1d1+m7A8Asrqk47VyR29o3n9nbtd1im/CzMMLR1u/SUdAb/ar5aa7By0QV+HuTBVMXtl8GGGzezraxXXMQ3+96bGOru6bAnNf7D608EUBgNXWKGW0nJ8BsOCtY4or1Ise5f+FKCBa2HtqBUwujWK0LqbBXMfThqVFO56CbgUNtAulwa0uYK2wkHM9WtiOecHkqRcj7UEAqH+ZwkVq5fS0ctzRcPxSNhtzC5yUc5NO03pFABQWRFc/w5jWC7oSpgr4TJoDLB0JdCfdBfH7VSbh0UPbSqnj5XvxK2aXP4P485IkSZIkSZIkSZIkSZIkSZIkSZIk8Tv/B3bBREdOWYS3AAAAAElFTkSuQmCC";
return pictureBase64;
```

---

### insertOoxml

**Kind:** `create`

Inserts OOXML at the specified location.

#### Signature

**Parameters:**
- `ooxml`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert formatted text with custom font size, color, line spacing, and paragraph spacing at the beginning of the current selection using OOXML.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('OOXML added to the beginning of the range.');
});

// Read "Create better add-ins for Word with Office Open XML" for guidance on working with OOXML.
// https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml
```

---

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a new paragraph with specified content immediately after the current text selection in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Paragraph added to the end of the range.');
});
```

---

### insertPictureFromBase64

**Kind:** `create`

Inserts a floating picture in front of text with its anchor at the beginning of the range.

#### Signature

**Parameters:**
- `base64EncodedImage`: `None` (required)
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a floating company logo image at the beginning of the selected text range

```typescript
await Word.run(async (context) => {
    // Get the selected range
    const range = context.document.getSelection();
    
    // Base64 encoded image string (example: a small PNG image)
    const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
    
    // Insert floating picture at the beginning of the range
    const picture = range.insertPictureFromBase64(base64Image, {
        width: 100,
        height: 100,
        left: 0,
        top: 0
    });
    
    await context.sync();
    
    console.log("Picture inserted successfully");
});
```

---

### insertTable

**Kind:** `create`

Inserts a table with the specified number of rows and columns.

#### Signature

**Parameters:**
- `rowCount`: `None` (required)
- `columnCount`: `None` (required)
- `insertLocation`: `None` (required)
- `values`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a 3x4 table with sample data at the end of the selected range

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    const tableData = [
        ["Product", "Q1", "Q2", "Q3"],
        ["Laptops", "150", "175", "200"],
        ["Monitors", "80", "95", "110"]
    ];
    
    const table = range.insertTable(3, 4, Word.InsertLocation.end, tableData);
    
    await context.sync();
});
```

---

### insertText

**Kind:** `create`

Inserts text at the specified location.

#### Signature

**Parameters:**
- `text`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert the text "New text inserted into the range." at the end of the currently selected range in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Text added to the end of the range.');
});
```

---

### insertTextBox

**Kind:** `create`

Inserts a floating text box in front of text with its anchor at the beginning of the range.

#### Signature

**Parameters:**
- `text`: `None` (required)
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a text box with placeholder text at the beginning of the current selection with specified dimensions and position.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a text box at the beginning of the selection.
  const range: Word.Range = context.document.getSelection();
  const insertShapeOptions: Word.InsertShapeOptions = {
    top: 0,
    left: 0,
    height: 100,
    width: 100
  };

  const newTextBox: Word.Shape = range.insertTextBox("placeholder text", insertShapeOptions);
  await context.sync();

  console.log("Inserted a text box at the beginning of the current selection.");
});
```

---

### intersectWith

Returns a new range as the intersection of this range with another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges aren't overlapped or adjacent.

#### Signature

**Parameters:**
- `range`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find and highlight the overlapping text between the first paragraph and the second paragraph in the document.

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    if (paragraphs.items.length >= 2) {
        const firstRange = paragraphs.items[0].getRange();
        const secondRange = paragraphs.items[1].getRange();
        
        try {
            // Get the intersection of the two ranges
            const intersectionRange = firstRange.intersectWith(secondRange);
            intersectionRange.font.highlightColor = "yellow";
            
            await context.sync();
            console.log("Intersection highlighted successfully");
        } catch (error) {
            console.log("No intersection found between the two paragraphs");
        }
    }
});
```

---

### intersectWithOrNullObject

Returns a new range as the intersection of this range with another range. This range isn't changed. If the two ranges aren't overlapped or adjacent, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Signature

**Parameters:**
- `range`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find and highlight the overlapping text between the first paragraph and the second paragraph in the document, or show a message if they don't overlap.

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    if (paragraphs.items.length >= 2) {
        const firstRange = paragraphs.items[0].getRange();
        const secondRange = paragraphs.items[1].getRange();
        
        const intersection = firstRange.intersectWithOrNullObject(secondRange);
        intersection.load("isNullObject, text");
        await context.sync();

        if (intersection.isNullObject) {
            console.log("The two paragraphs do not overlap or are not adjacent.");
        } else {
            intersection.font.highlightColor = "yellow";
            console.log("Highlighted overlapping text: " + intersection.text);
        }
    }
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Get and display the text content of the currently selected range in a Word document

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Load the text property of the range
    range.load("text");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded property
    console.log("Selected text: " + range.text);
});
```

---

### removeHighlight

**Kind:** `delete`

Removes the highlight added by the Highlight function if any.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Remove highlighting from all highlighted text in the first paragraph of the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    // Remove any highlighting from the paragraph
    range.removeHighlight();
    
    await context.sync();
});
```

---

### search

Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `None` (required)
- `searchOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find all occurrences of the word "TODO" in a selected range and highlight them in yellow

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    
    const searchResults = range.search("TODO", { matchCase: false });
    searchResults.load("items");
    
    await context.sync();
    
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### select

Selects and navigates the Word UI to the range.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Insert HTML text at the beginning of the current selection and select the inserted content.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Selected the range.');
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Set multiple formatting properties on the first paragraph's range at once, including font color, size, and bold styling.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const range = firstParagraph.getRange();
    
    // Set multiple properties at once using the set() method
    range.set({
        font: {
            color: "#FF0000",
            size: 16,
            bold: true
        }
    });
    
    await context.sync();
});
```

---

### split

Splits the range into child ranges by using delimiters.

#### Signature

**Parameters:**
- `delimiters`: `None` (required)
- `multiParagraphs`: `None` (required)
- `trimDelimiters`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Split a text range containing comma-separated values into individual ranges and highlight each segment with alternating colors

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's range
    const paragraph = context.document.body.paragraphs.getFirst();
    const range = paragraph.getRange();
    range.load("text");
    
    await context.sync();
    
    // Split the range by commas
    const childRanges = range.split([","], false, true, true);
    childRanges.load("items");
    
    await context.sync();
    
    // Highlight each segment with alternating colors
    for (let i = 0; i < childRanges.items.length; i++) {
        childRanges.items[i].font.highlightColor = i % 2 === 0 ? "yellow" : "lightblue";
    }
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Range` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize a range object to JSON format to log or store its properties after loading specific properties like text and font name.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's range
    const paragraph = context.document.body.paragraphs.getFirst();
    const range = paragraph.getRange();
    
    // Load properties you want to serialize
    range.load("text, font/name, font/size");
    
    await context.sync();
    
    // Convert the range to a plain JavaScript object
    const rangeData = range.toJSON();
    
    // Now you can use JSON.stringify or access properties as plain data
    console.log(JSON.stringify(rangeData, null, 2));
    console.log("Range text:", rangeData.text);
    console.log("Font name:", rangeData.font?.name);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get a range from the current selection, track it across multiple sync calls, and apply formatting changes while maintaining the reference to the same range object.

```typescript
await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    
    // Track the range to use it across multiple sync calls
    range.track();
    
    // First sync: apply bold formatting
    range.font.bold = true;
    await context.sync();
    
    // Second sync: apply color (range reference still valid because it's tracked)
    range.font.color = "blue";
    await context.sync();
    
    // Third sync: apply highlight
    range.font.highlightColor = "yellow";
    await context.sync();
    
    // Clean up: untrack when done
    range.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Search for text in a document, highlight all instances, then untrack the range objects to free memory after processing.

```typescript
await Word.run(async (context) => {
    // Search for all instances of "TODO" in the document
    const searchResults = context.document.body.search("TODO", { matchCase: false });
    
    // Track the search results to work with them
    context.load(searchResults, "items");
    await context.sync();
    
    // Highlight each found range
    for (let i = 0; i < searchResults.items.length; i++) {
        const range = searchResults.items[i];
        range.font.highlightColor = "yellow";
    }
    
    await context.sync();
    
    // Untrack all range objects to free memory
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].untrack();
    }
    
    await context.sync();
    
    console.log(`Highlighted ${searchResults.items.length} instances and freed memory.`);
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
