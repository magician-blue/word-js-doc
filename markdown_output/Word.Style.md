# Word.Style

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a style in a Word document.

## Class Examples

```typescript
// Link to full sample: // Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

## Properties

### automaticallyUpdate

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the style is automatically redefined based on the selection.

#### Examples

**Example**: Disable automatic style updates for the "Heading 1" style to prevent it from changing when users modify formatted text

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Disable automatic updates for this style
    heading1Style.automaticallyUpdate = false;
    
    await context.sync();
    
    console.log("Automatic updates disabled for Heading 1 style");
});
```

---

### baseStyle

**Type:** `string`

**Since:** WordApi 1.5

Specifies the name of an existing style to use as the base formatting of another style.

#### Examples

**Example**: Set a custom style to inherit formatting from the built-in "Heading 1" style as its base

```typescript
await Word.run(async (context) => {
    // Get or create a custom style
    const customStyle = context.document.getStyles().getByNameOrNullObject("MyCustomStyle");
    await context.sync();

    let style: Word.Style;
    if (customStyle.isNullObject) {
        style = context.document.getStyles().add("MyCustomStyle", Word.StyleType.paragraph);
    } else {
        style = customStyle;
    }

    // Set "Heading 1" as the base style
    style.baseStyle = "Heading 1";

    await context.sync();
    console.log("Custom style now inherits from Heading 1");
});
```

---

### borders

**Type:** `Word.BorderCollection`

**Since:** WordApiDesktop 1.1

Specifies a BorderCollection object that represents all the borders for the specified style.

#### Examples

**Example**: Update the outside border properties of a specified style to use dashed green borders with 0.25 point width.

```typescript
// Link to full sample: // Updates border properties (e.g., type, width, color) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update border properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const borders: Word.BorderCollection = style.borders;
    borders.load("items");
    await context.sync();

    borders.outsideBorderType = Word.BorderType.dashed;
    borders.outsideBorderWidth = Word.BorderWidth.pt025;
    borders.outsideBorderColor = "green";
    console.log("Updated outside borders.");
  }
});
```

---

### builtIn

**Type:** `boolean`

**Since:** WordApi 1.5

Gets whether the specified style is a built-in style.

#### Examples

**Example**: Check if a paragraph's style is a built-in Word style and display the result in the console

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const style = paragraph.style;
    
    // Load the style object with the builtIn property
    context.load(style, "builtIn, name");
    
    await context.sync();
    
    console.log(`Style "${style.name}" is ${style.builtIn ? "a built-in" : "a custom"} style`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Style object to load and read the style's name property

```typescript
await Word.run(async (context) => {
    // Get a style from the document
    const style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Access the request context associated with the style object
    const styleContext = style.context;
    
    // Use the context to load properties
    style.load("nameLocal");
    
    await styleContext.sync();
    
    if (!style.isNullObject) {
        console.log("Style name: " + style.nameLocal);
    }
});
```

---

### description

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the description of the specified style.

#### Examples

**Example**: Display the description of the "Heading 1" style in the console to understand its purpose and usage.

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style from the document
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Load the description property
    heading1Style.load("description");
    
    await context.sync();
    
    // Display the style description
    if (!heading1Style.isNullObject) {
        console.log("Heading 1 style description: " + heading1Style.description);
    } else {
        console.log("Heading 1 style not found");
    }
});
```

---

### font

**Type:** `Word.Font`

**Since:** WordApi 1.5

Gets a font object that represents the character formatting of the specified style.

#### Examples

**Example**: Update the font color to red and font size to 20 points for a specified style in the document.

```typescript
// Link to full sample: // Updates font properties (e.g., color, size) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update font properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const font: Word.Font = style.font;
    font.color = "#FF0000";
    font.size = 20;
    console.log(`Successfully updated font properties of the '${styleName}' style.`);
  }
});
```

---

### frame

**Type:** `Word.Frame`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Frame object that represents the frame formatting for the style.

#### Examples

**Example**: Get the frame settings of a paragraph style and display whether the style has frame formatting enabled.

```typescript
await Word.run(async (context) => {
    // Get a style by name
    const style = context.document.getStyles().getByNameOrNullObject("MyCustomStyle");
    
    // Get the frame object for this style
    const frame = style.frame;
    
    // Load properties to check frame settings
    frame.load("width, height, horizontalPosition");
    
    await context.sync();
    
    if (!style.isNullObject) {
        console.log(`Frame width: ${frame.width}`);
        console.log(`Frame height: ${frame.height}`);
        console.log(`Frame horizontal position: ${frame.horizontalPosition}`);
    } else {
        console.log("Style not found");
    }
});
```

---

### hasProofing

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the spelling and grammar checker ignores text formatted with this style.

#### Examples

**Example**: Disable spell-checking and grammar-checking for text formatted with the "Code" style by setting hasProofing to true

```typescript
await Word.run(async (context) => {
    // Get the "Code" style from the document
    const codeStyle = context.document.getStyles().getByNameOrNullObject("Code");
    
    // Load the style properties
    await context.sync();
    
    // Check if the style exists
    if (!codeStyle.isNullObject) {
        // Set hasProofing to true to ignore spelling and grammar checking
        codeStyle.hasProofing = true;
        
        await context.sync();
        console.log("Proofing disabled for Code style");
    } else {
        console.log("Code style not found");
    }
});
```

---

### inUse

**Type:** `boolean`

**Since:** WordApi 1.5

Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.

#### Examples

**Example**: Check if a specific style has been used or modified in the document and display a message to the user

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    heading1Style.load("inUse");
    
    await context.sync();
    
    if (!heading1Style.isNullObject) {
        if (heading1Style.inUse) {
            console.log("Heading 1 style is being used or has been modified in this document.");
        } else {
            console.log("Heading 1 style has not been used or modified in this document.");
        }
    }
});
```

---

### languageId

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a LanguageId value that represents the language for the style.

#### Examples

**Example**: Set the language of the "Heading 1" style to French

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Set the language to French
    heading1Style.languageId = "French";
    
    await context.sync();
    
    console.log("Language set to French for Heading 1 style");
});
```

---

### languageIdFarEast

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies an East Asian language for the style.

#### Examples

**Example**: Set the Far East language to Japanese for the "Heading 1" style to ensure proper text rendering for Japanese characters

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Set the Far East language to Japanese
    heading1Style.languageIdFarEast = "Japanese";
    
    await context.sync();
    
    console.log("Far East language set to Japanese for Heading 1 style");
});
```

---

### linked

**Type:** `boolean`

**Since:** WordApi 1.5

Gets whether a style is a linked style that can be used for both paragraph and character formatting.

#### Examples

**Example**: Check if the "Heading 1" style is a linked style that can be used for both paragraph and character formatting, and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Load the linked property
    heading1Style.load("linked");
    
    await context.sync();
    
    // Check if the style is linked
    if (!heading1Style.isNullObject) {
        console.log(`Heading 1 is a linked style: ${heading1Style.linked}`);
    } else {
        console.log("Heading 1 style not found");
    }
});
```

---

### linkStyle

**Type:** `Word.Style`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a link between a paragraph and a character style.

#### Examples

**Example**: Link a paragraph style named "CustomHeading" to a character style named "CustomHeadingChar" so that character formatting can be applied independently within paragraphs using the heading style.

```typescript
await Word.run(async (context) => {
    // Get the paragraph style and character style
    const paragraphStyle = context.document.getStyles().getByNameOrNullObject("CustomHeading");
    const characterStyle = context.document.getStyles().getByNameOrNullObject("CustomHeadingChar");
    
    // Load the styles
    paragraphStyle.load("name");
    characterStyle.load("name");
    
    await context.sync();
    
    // Check if both styles exist
    if (!paragraphStyle.isNullObject && !characterStyle.isNullObject) {
        // Link the character style to the paragraph style
        paragraphStyle.linkStyle = characterStyle;
        
        await context.sync();
        console.log("Successfully linked CustomHeadingChar to CustomHeading style");
    } else {
        console.log("One or both styles do not exist");
    }
});
```

---

### listLevelNumber

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the list level for the style.

#### Examples

**Example**: Get the list level number of a paragraph's style to determine its position in the outline hierarchy

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const style = paragraph.style;
    
    // Load the style object
    context.load(style);
    await context.sync();
    
    // Get the style by name and load its list level number
    const styleObject = context.document.getStyles().getByNameOrNullObject(style);
    styleObject.load("listLevelNumber");
    await context.sync();
    
    if (!styleObject.isNullObject) {
        console.log(`List level number: ${styleObject.listLevelNumber}`);
        // Returns 0-8 for heading levels, or 0 if not a list style
    }
});
```

---

### listTemplate

**Type:** `Word.ListTemplate`

**Since:** WordApiDesktop 1.1

Gets a ListTemplate object that represents the list formatting for the specified Style object.

#### Examples

**Example**: Retrieve and display the properties and list levels of a specified list style from the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

---

### locked

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the style cannot be changed or edited.

#### Examples

**Example**: Check if the "Heading 1" style is locked and display the result in the console

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Load the locked property
    heading1Style.load("locked");
    
    await context.sync();
    
    // Check if the style exists and display its locked status
    if (!heading1Style.isNullObject) {
        console.log(`Heading 1 style is ${heading1Style.locked ? "locked" : "unlocked"}`);
    } else {
        console.log("Heading 1 style not found");
    }
});
```

---

### nameLocal

**Type:** `string`

**Since:** WordApi 1.5

Gets the name of a style in the language of the user.

#### Examples

**Example**: Apply a user-specified paragraph style to the first paragraph of the document body after validating the style exists and is of paragraph type.

```typescript
// Link to full sample: // Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

---

### nextParagraphStyle

**Type:** `string`

**Since:** WordApi 1.5

Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.

#### Examples

**Example**: Set the "Heading 1" style so that when a user presses Enter after a Heading 1 paragraph, the next paragraph automatically uses the "Normal" style instead of continuing with Heading 1.

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Set the next paragraph style to "Normal"
    heading1Style.nextParagraphStyle = "Normal";
    
    await context.sync();
    
    console.log("Heading 1 will now automatically switch to Normal style for the next paragraph");
});
```

---

### noSpaceBetweenParagraphsOfSameStyle

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove spacing between paragraphs that are formatted using the same style.

#### Examples

**Example**: Remove spacing between paragraphs that use the "Heading 1" style to create a more compact appearance

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Load the style
    heading1Style.load("noSpaceBetweenParagraphsOfSameStyle");
    await context.sync();
    
    // Remove spacing between paragraphs of the same style
    if (!heading1Style.isNullObject) {
        heading1Style.noSpaceBetweenParagraphsOfSameStyle = true;
        await context.sync();
        
        console.log("Spacing removed between consecutive Heading 1 paragraphs");
    }
});
```

---

### paragraphFormat

**Type:** `Word.ParagraphFormat`

**Since:** WordApi 1.5

Gets a ParagraphFormat object that represents the paragraph settings for the specified style.

#### Examples

**Example**: Update a Word style's paragraph format by setting its left indent to 30 points and alignment to centered.

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

### priority

**Type:** `number`

**Since:** WordApi 1.5

Specifies the priority.

#### Examples

**Example**: Set the priority of the "Heading 1" style to 10 to control its position in the style gallery

```typescript
await Word.run(async (context) => {
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    heading1Style.load("priority");
    await context.sync();
    
    if (!heading1Style.isNullObject) {
        heading1Style.priority = 10;
        await context.sync();
    }
});
```

---

### quickStyle

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether the style corresponds to an available quick style.

#### Examples

**Example**: Check if a paragraph's style is a quick style and display an alert with the result

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const style = paragraph.style;
    
    // Load the style object to access its properties
    context.load(style, "quickStyle");
    
    await context.sync();
    
    if (style.quickStyle) {
        console.log("This paragraph uses a quick style.");
    } else {
        console.log("This paragraph does not use a quick style.");
    }
});
```

---

### shading

**Type:** `Word.Shading`

**Since:** WordApi 1.6

Gets a Shading object that represents the shading for the specified style. Not applicable to List style.

#### Examples

**Example**: Update the shading properties of a specified style by setting its background pattern color to blue, foreground pattern color to yellow, and texture to dark trellis.

```typescript
// Link to full sample: // Updates shading properties (e.g., texture, pattern colors) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update shading properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const shading: Word.Shading = style.shading;
    shading.load();
    await context.sync();

    shading.backgroundPatternColor = "blue";
    shading.foregroundPatternColor = "yellow";
    shading.texture = Word.ShadingTextureType.darkTrellis;

    console.log("Updated shading.");
  }
});
```

---

### tableStyle

**Type:** `Word.TableStyle`

**Since:** WordApi 1.6

Gets a TableStyle object representing Style properties that can be applied to a table.

#### Examples

**Example**: Apply table-specific formatting to a style by setting the table style's row banding and header row properties

```typescript
await Word.run(async (context) => {
    // Get a style from the document
    const style = context.document.getStyles().getByNameOrNullObject("MyTableStyle");
    style.load("type");
    
    await context.sync();
    
    if (!style.isNullObject && style.type === Word.StyleType.table) {
        // Access the tableStyle property to configure table-specific formatting
        const tableStyle = style.tableStyle;
        
        // Enable row banding and configure header row
        tableStyle.allowBreakAcrossPage = false;
        
        await context.sync();
        console.log("Table style properties configured successfully");
    }
});
```

---

### type

**Type:** `Word.StyleType | "Character" | "List" | "Paragraph" | "Table"`

**Since:** WordApi 1.5

Gets the style type.

#### Examples

**Example**: Check if a specific style is a paragraph style before applying it to selected text

```typescript
await Word.run(async (context) => {
    // Get the style named "Heading 1"
    const style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    style.load("type");
    
    await context.sync();
    
    if (!style.isNullObject) {
        // Check if the style is a paragraph style
        if (style.type === Word.StyleType.paragraph || style.type === "Paragraph") {
            console.log("Heading 1 is a paragraph style");
            
            // Safe to apply to the current paragraph
            const paragraph = context.document.getSelection().paragraphs.getFirst();
            paragraph.style = "Heading 1";
        } else {
            console.log(`Heading 1 is a ${style.type} style, not a paragraph style`);
        }
    }
    
    await context.sync();
});
```

---

### unhideWhenUsed

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.

#### Examples

**Example**: Configure a custom style to automatically appear in the recommended styles gallery after it has been used in the document

```typescript
await Word.run(async (context) => {
    // Get the style by name
    const style = context.document.getStyles().getByNameOrNullObject("MyCustomStyle");
    
    // Load the style properties
    style.load("unhideWhenUsed");
    await context.sync();
    
    // Check if style exists
    if (!style.isNullObject) {
        // Set the style to appear in recommended styles after use
        style.unhideWhenUsed = true;
        
        await context.sync();
        console.log("Style will be visible in recommended styles after use");
    }
});
```

---

### visibility

**Type:** `boolean`

**Since:** WordApi 1.5

Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

#### Examples

**Example**: Hide a custom style named "CustomHeading" from appearing in the Styles gallery and Styles task pane

```typescript
await Word.run(async (context) => {
    const style = context.document.getStyles().getByNameOrNullObject("CustomHeading");
    style.load("visibility");
    
    await context.sync();
    
    if (!style.isNullObject) {
        style.visibility = false;
        await context.sync();
        console.log("Style 'CustomHeading' is now hidden from the Styles gallery");
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the style.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a custom style from the Word document by retrieving it by name and removing it if it exists.

```typescript
// Link to full sample: // Deletes the custom style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to delete.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.delete();
    console.log(`Successfully deleted custom style '${styleName}'.`);
  }
});
```

---

### linkToListTemplate

Links this style to a list template so that the style's formatting can be applied to lists.

#### Signature

**Parameters:**
- `listTemplate`: `Word.ListTemplate` (required)
  A ListTemplate to link to the style.

**Returns:** `void`

#### Examples

**Example**: Create a custom list style and link it to a list template to apply numbered formatting to paragraphs

```typescript
await Word.run(async (context) => {
    // Get or create a list template
    const listTemplate = context.document.body.lists.getFirst().listTemplate;
    
    // Get or create a custom style
    const customStyle = context.document.getStyles().getByNameOrNullObject("MyListStyle");
    await context.sync();
    
    let style: Word.Style;
    if (customStyle.isNullObject) {
        style = context.document.addStyle("MyListStyle", "Paragraph");
    } else {
        style = customStyle;
    }
    
    // Link the style to the list template
    style.linkToListTemplate(listTemplate);
    
    // Apply the style to a paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.style = "MyListStyle";
    
    await context.sync();
    console.log("Style linked to list template and applied to paragraph");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.StyleLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Style`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Style`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{
            select?: string;
            expand?: string;
        }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Style`

#### Examples

**Example**: Load and display the name and font properties of the "Heading 1" style from the document.

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Load specific properties of the style
    heading1Style.load("nameLocal, font/name, font/size, font/color");
    
    // Sync to execute the load command
    await context.sync();
    
    // Check if style exists and display properties
    if (!heading1Style.isNullObject) {
        console.log(`Style Name: ${heading1Style.nameLocal}`);
        console.log(`Font Name: ${heading1Style.font.name}`);
        console.log(`Font Size: ${heading1Style.font.size}`);
        console.log(`Font Color: ${heading1Style.font.color}`);
    } else {
        console.log("Heading 1 style not found");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.StyleUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Style` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of the "Heading 1" style to change its font name, size, and color

```typescript
await Word.run(async (context) => {
    // Get the "Heading 1" style
    const heading1Style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    
    // Set multiple properties at once
    heading1Style.set({
        font: {
            name: "Arial",
            size: 16,
            color: "#0066CC"
        }
    });
    
    await context.sync();
    console.log("Heading 1 style updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify().

#### Signature

**Returns:** `Word.Interfaces.StyleData`

#### Examples

**Example**: Retrieve a paragraph's style properties and log them as a JSON string to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const style = paragraph.style;
    
    // Load the style properties
    style.load();
    
    // Sync to get the values
    await context.sync();
    
    // Convert the style object to JSON and log it
    const styleJSON = style.toJSON();
    console.log("Style properties as JSON:", JSON.stringify(styleJSON, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document.

#### Signature

**Returns:** `Word.Style`

#### Examples

**Example**: Track a custom style object to automatically maintain its reference while modifying multiple paragraphs in the document, ensuring the style properties remain accessible even as the document structure changes.

```typescript
await Word.run(async (context) => {
    // Get a custom style from the document
    const style = context.document.getStyles().getByNameOrNullObject("Heading1");
    
    // Track the style object for automatic adjustment
    style.track();
    
    // Load style properties
    style.load("font/bold,font/size");
    
    // Make changes to the document that might affect object references
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    
    await context.sync();
    
    // Add new paragraphs which changes the document structure
    context.document.body.insertParagraph("New paragraph 1", "Start");
    context.document.body.insertParagraph("New paragraph 2", "Start");
    
    await context.sync();
    
    // The tracked style object is still valid and accessible
    console.log(`Style font size: ${style.font.size}`);
    console.log(`Style font bold: ${style.font.bold}`);
    
    // Untrack when done to release memory
    style.untrack();
    
    await context.sync();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked.

#### Signature

**Returns:** `Word.Style`

#### Examples

**Example**: Load a style object to check its properties, then release it from memory tracking to optimize performance when the style is no longer needed.

```typescript
await Word.run(async (context) => {
    // Get a style and track it
    const style = context.document.getStyles().getByNameOrNullObject("Heading 1");
    style.track();
    
    // Load and use the style
    style.load("font/size");
    await context.sync();
    
    if (!style.isNullObject) {
        console.log(`Heading 1 font size: ${style.font.size}`);
        
        // Release the style from tracking when done
        style.untrack();
    }
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml
