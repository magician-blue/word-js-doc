# Word.BorderUniversal

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `ClientObject`

## Description

Represents the BorderUniversal object, which manages borders for a range, paragraph, table, or frame.

## Properties

### artStyle

**Type:** `Word.PageBorderArt | "Apples" | "MapleMuffins" | "CakeSlice" | "CandyCorn" | "IceCreamCones" | "ChampagneBottle" | "PartyGlass" | "ChristmasTree" | "Trees" | "PalmsColor" | "Balloons3Colors" | "BalloonsHotAir" | "PartyFavor" | "ConfettiStreamers" | "Hearts" | "HeartBalloon" | "Stars3D" | "StarsShadowed" | "Stars" | "Sun" | "Earth2" | "Earth1" | "PeopleHats" | "Sombrero" | "Pencils" | "Packages" | "Clocks" | "Firecrackers" | "Rings" | "MapPins" | "Confetti" | "CreaturesButterfly" | "CreaturesLadyBug" | "CreaturesFish" | "BirdsFlight" | "ScaredCat" | "Bats" | "FlowersRoses" | "FlowersRedRose" | "Poinsettias" | "Holly" | "FlowersTiny" | "FlowersPansy" | "FlowersModern2" | "FlowersModern1" | "WhiteFlowers" | "Vine" | "FlowersDaisies" | "FlowersBlockPrint" | "DecoArchColor" | "Fans" | "Film" | "Lightning1" | "Compass" | "DoubleD" | "ClassicalWave" | "ShadowedSquares" | "TwistedLines1" | "Waveline" | "Quadrants" | "CheckedBarColor" | "Swirligig" | "PushPinNote1" | "PushPinNote2" | "Pumpkin1" | "EggsBlack" | "Cup" | "HeartGray" | "GingerbreadMan" | "BabyPacifier" | "BabyRattle" | "Cabins" | "HouseFunky" | "StarsBlack" | "Snowflakes" | "SnowflakeFancy" | "Skyrocket" | "Seattle" | "MusicNotes" | "PalmsBlack" | "MapleLeaf" | "PaperClips" | "ShorebirdTracks" | "People" | "PeopleWaving" | "EclipsingSquares2" | "Hypnotic" | "DiamondsGray" | "DecoArch" | "DecoBlocks" | "CirclesLines" | "Papyrus" | "Woodwork" | "WeavingBraid" | "WeavingRibbon" | "WeavingAngles" | "ArchedScallops" | "Safari" | "CelticKnotwork" | "CrazyMaze" | "EclipsingSquares1" | "Birds" | "FlowersTeacup" | "Northwest" | "Southwest" | "Tribal6" | "Tribal4" | "Tribal3" | "Tribal2" | "Tribal5" | "XIllusions" | "ZanyTriangles" | "Pyramids" | "PyramidsAbove" | "ConfettiGrays" | "ConfettiOutline" | "ConfettiWhite" | "Mosaic" | "Lightning2" | "HeebieJeebies" | "LightBulb" | "Gradient" | "TriangleParty" | "TwistedLines2" | "Moons" | "Ovals" | "DoubleDiamonds" | "ChainLink" | "Triangles" | "Tribal1" | "MarqueeToothed" | "SharksTeeth" | "Sawtooth" | "SawtoothGray" | "PostageStamp" | "WeavingStrips" | "ZigZag" | "CrossStitch" | "Gems" | "CirclesRectangles" | "CornerTriangles" | "CreaturesInsects" | "ZigZagStitch" | "Checkered" | "CheckedBarBlack" | "Marquee" | "BasicWhiteDots" | "BasicWideMidline" | "BasicWideOutline" | "BasicWideInline" | "BasicThinLines" | "BasicWhiteDashes" | "BasicWhiteSquares" | "BasicBlackSquares" | "BasicBlackDashes" | "BasicBlackDots" | "StarsTop" | "CertificateBanner" | "Handmade1" | "Handmade2" | "TornPaper" | "TornPaperBlack" | "CouponCutoutDashes" | "CouponCutoutDots"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the graphical page-border design for the document.

#### Examples

**Example**: Apply a decorative stars border design to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Set the art style to a decorative stars design
    paragraph.border.artStyle = "Stars3D";
    
    await context.sync();
});
```

---

### artWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width (in points) of the graphical page border specified in the artStyle property.

#### Examples

**Example**: Set the graphical page border width to 24 points for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Set the art width to 24 points
    border.artWidth = 24;
    
    await context.sync();
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the border color of the first paragraph to red using the hex color format

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    border.color = "#FF0000";
    
    await context.sync();
});
```

---

### colorIndex

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color for the BorderUniversal or Word.Font object.

#### Examples

**Example**: Set the border color of the first paragraph to red using the colorIndex property

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Set the border color to red
    border.colorIndex = "Red";
    border.visible = true;
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BorderUniversal object to synchronize border properties with the Office host application

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border object for the paragraph
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Access the request context from the border object
    const borderContext = border.context;
    
    // Use the context to load and sync border properties
    border.load("type,color,width");
    await borderContext.sync();
    
    console.log(`Border type: ${border.type}`);
    console.log(`Border color: ${border.color}`);
    console.log(`Border width: ${border.width}`);
});
```

---

### inside

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns true if an inside border can be applied to the specified object.

#### Examples

**Example**: Check if an inside border can be applied to a table and display the result in the console

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    const borders = table.getBorder(Word.BorderLocation.inside);
    borders.load("inside");
    
    await context.sync();
    
    console.log(`Can apply inside border: ${borders.inside}`);
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the border is visible.

#### Examples

**Example**: Make the border of the first paragraph invisible

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const border = firstParagraph.getBorder(Word.BorderLocation.top);
    
    border.isVisible = false;
    
    await context.sync();
});
```

---

### lineStyle

**Type:** `Word.BorderLineStyle | "None" | "Single" | "Dot" | "DashSmallGap" | "DashLargeGap" | "DashDot" | "DashDotDot" | "Double" | "Triple" | "ThinThickSmallGap" | "ThickThinSmallGap" | "ThinThickThinSmallGap" | "ThinThickMedGap" | "ThickThinMedGap" | "ThinThickThinMedGap" | "ThinThickLargeGap" | "ThickThinLargeGap" | "ThinThickThinLargeGap" | "SingleWavy" | "DoubleWavy" | "DashDotStroked" | "Emboss3D" | "Engrave3D" | "Outset" | "Inset"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the line style of the border.

#### Examples

**Example**: Set a paragraph's bottom border to use a double-line style

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.bottom);
    
    border.lineStyle = "Double";
    border.visible = true;
    
    await context.sync();
});
```

---

### lineWidth

**Type:** `Word.LineWidth | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the line width of an object's border.

#### Examples

**Example**: Set the border line width of the first paragraph to 2.25 points

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    border.lineWidth = Word.LineWidth.pt225;
    
    await context.sync();
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
  - `options`: `Word.Interfaces.BorderUniversalLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BorderUniversal`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BorderUniversal`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BorderUniversal`

#### Examples

**Example**: Get and display the border color of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const border = firstParagraph.getBorder(Word.BorderLocation.top);
    
    // Load the color property of the border
    border.load("color");
    
    await context.sync();
    
    console.log("Border color: " + border.color);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BorderUniversalUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.BorderUniversal` (required)

  **Returns:** `void`

#### Examples

**Example**: Set multiple border properties at once for the first paragraph, including color, line style, and width

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const border = firstParagraph.getBorder(Word.BorderLocation.top);
    
    // Set multiple border properties at once
    border.set({
        color: "#FF0000",
        lineStyle: Word.BorderLineStyle.single,
        width: 3
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderUniversalData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BorderUniversalData`

#### Examples

**Example**: Get a paragraph's border properties as a plain JavaScript object and log it to the console for inspection or serialization.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border object for the paragraph
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Load border properties
    border.load("type, color, width, visible");
    
    await context.sync();
    
    // Convert the border object to a plain JavaScript object
    const borderData = border.toJSON();
    
    // Now you can use the plain object for logging, serialization, etc.
    console.log("Border properties:", borderData);
    console.log("Border color:", borderData.color);
    console.log("Border width:", borderData.width);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BorderUniversal`

#### Examples

**Example**: Apply a border to a paragraph and track it across multiple sync calls to modify its properties without getting an InvalidObjectPath error

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Track the border object for use across multiple sync calls
    border.track();
    
    await context.sync();
    
    // Now we can safely modify the border properties after sync
    border.type = Word.BorderType.single;
    border.color = "#FF0000";
    border.width = 2;
    
    await context.sync();
    
    // Can continue to work with the tracked border object
    border.width = 4;
    
    await context.sync();
    
    // Untrack when done to free up memory
    border.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BorderUniversal`

#### Examples

**Example**: Apply a border to a paragraph, then untrack the border object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border object and track it
    const border = paragraph.getBorder(Word.BorderLocation.top);
    border.track();
    border.load("type");
    
    await context.sync();
    
    // Apply border settings
    border.type = Word.BorderType.single;
    border.color = "#0000FF";
    border.width = 2;
    
    await context.sync();
    
    // Untrack the border object to release memory
    border.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://learn.microsoft.com/en-us/javascript/api/word/word.pageborderart
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext
- https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderlinestyle
- https://learn.microsoft.com/en-us/javascript/api/word/word.linewidth
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversalloadoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversalupdatedata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversaldata
