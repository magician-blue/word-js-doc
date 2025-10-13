# Word.Interfaces.BorderUniversalData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling borderUniversal.toJSON().

## Properties

- [artStyle](#artstyle)
  - Specifies the graphical page-border design for the document.
- [artWidth](#artwidth)
  - Specifies the width (in points) of the graphical page border specified in the artStyle property.
- [color](#color)
  - Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.
- [colorIndex](#colorindex)
  - Specifies the color for the BorderUniversal or [Word.Font](/en-us/javascript/api/word/word.font) object.
- [inside](#inside)
  - Returns true if an inside border can be applied to the specified object.
- [isVisible](#isvisible)
  - Specifies whether the border is visible.
- [lineStyle](#linestyle)
  - Specifies the line style of the border.
- [lineWidth](#linewidth)
  - Specifies the line width of an object's border.

## Property Details

### artStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the graphical page-border design for the document.

```typescript
artStyle?: Word.PageBorderArt | "Apples" | "MapleMuffins" | "CakeSlice" | "CandyCorn" | "IceCreamCones" | "ChampagneBottle" | "PartyGlass" | "ChristmasTree" | "Trees" | "PalmsColor" | "Balloons3Colors" | "BalloonsHotAir" | "PartyFavor" | "ConfettiStreamers" | "Hearts" | "HeartBalloon" | "Stars3D" | "StarsShadowed" | "Stars" | "Sun" | "Earth2" | "Earth1" | "PeopleHats" | "Sombrero" | "Pencils" | "Packages" | "Clocks" | "Firecrackers" | "Rings" | "MapPins" | "Confetti" | "CreaturesButterfly" | "CreaturesLadyBug" | "CreaturesFish" | "BirdsFlight" | "ScaredCat" | "Bats" | "FlowersRoses" | "FlowersRedRose" | "Poinsettias" | "Holly" | "FlowersTiny" | "FlowersPansy" | "FlowersModern2" | "FlowersModern1" | "WhiteFlowers" | "Vine" | "FlowersDaisies" | "FlowersBlockPrint" | "DecoArchColor" | "Fans" | "Film" | "Lightning1" | "Compass" | "DoubleD" | "ClassicalWave" | "ShadowedSquares" | "TwistedLines1" | "Waveline" | "Quadrants" | "CheckedBarColor" | "Swirligig" | "PushPinNote1" | "PushPinNote2" | "Pumpkin1" | "EggsBlack" | "Cup" | "HeartGray" | "GingerbreadMan" | "BabyPacifier" | "BabyRattle" | "Cabins" | "HouseFunky" | "StarsBlack" | "Snowflakes" | "SnowflakeFancy" | "Skyrocket" | "Seattle" | "MusicNotes" | "PalmsBlack" | "MapleLeaf" | "PaperClips" | "ShorebirdTracks" | "People" | "PeopleWaving" | "EclipsingSquares2" | "Hypnotic" | "DiamondsGray" | "DecoArch" | "DecoBlocks" | "CirclesLines" | "Papyrus" | "Woodwork" | "WeavingBraid" | "WeavingRibbon" | "WeavingAngles" | "ArchedScallops" | "Safari" | "CelticKnotwork" | "CrazyMaze" | "EclipsingSquares1" | "Birds" | "FlowersTeacup" | "Northwest" | "Southwest" | "Tribal6" | "Tribal4" | "Tribal3" | "Tribal2" | "Tribal5" | "XIllusions" | "ZanyTriangles" | "Pyramids" | "PyramidsAbove" | "ConfettiGrays" | "ConfettiOutline" | "ConfettiWhite" | "Mosaic" | "Lightning2" | "HeebieJeebies" | "LightBulb" | "Gradient" | "TriangleParty" | "TwistedLines2" | "Moons" | "Ovals" | "DoubleDiamonds" | "ChainLink" | "Triangles" | "Tribal1" | "MarqueeToothed" | "SharksTeeth" | "Sawtooth" | "SawtoothGray" | "PostageStamp" | "WeavingStrips" | "ZigZag" | "CrossStitch" | "Gems" | "CirclesRectangles" | "CornerTriangles" | "CreaturesInsects" | "ZigZagStitch" | "Checkered" | "CheckedBarBlack" | "Marquee" | "BasicWhiteDots" | "BasicWideMidline" | "BasicWideOutline" | "BasicWideInline" | "BasicThinLines" | "BasicWhiteDashes" | "BasicWhiteSquares" | "BasicBlackSquares" | "BasicBlackDashes" | "BasicBlackDots" | "StarsTop" | "CertificateBanner" | "Handmade1" | "Handmade2" | "TornPaper" | "TornPaperBlack" | "CouponCutoutDashes" | "CouponCutoutDots";
```

Property Value
[Word.PageBorderArt](/en-us/javascript/api/word/word.pageborderart) | "Apples" | "MapleMuffins" | "CakeSlice" | "CandyCorn" | "IceCreamCones" | "ChampagneBottle" | "PartyGlass" | "ChristmasTree" | "Trees" | "PalmsColor" | "Balloons3Colors" | "BalloonsHotAir" | "PartyFavor" | "ConfettiStreamers" | "Hearts" | "HeartBalloon" | "Stars3D" | "StarsShadowed" | "Stars" | "Sun" | "Earth2" | "Earth1" | "PeopleHats" | "Sombrero" | "Pencils" | "Packages" | "Clocks" | "Firecrackers" | "Rings" | "MapPins" | "Confetti" | "CreaturesButterfly" | "CreaturesLadyBug" | "CreaturesFish" | "BirdsFlight" | "ScaredCat" | "Bats" | "FlowersRoses" | "FlowersRedRose" | "Poinsettias" | "Holly" | "FlowersTiny" | "FlowersPansy" | "FlowersModern2" | "FlowersModern1" | "WhiteFlowers" | "Vine" | "FlowersDaisies" | "FlowersBlockPrint" | "DecoArchColor" | "Fans" | "Film" | "Lightning1" | "Compass" | "DoubleD" | "ClassicalWave" | "ShadowedSquares" | "TwistedLines1" | "Waveline" | "Quadrants" | "CheckedBarColor" | "Swirligig" | "PushPinNote1" | "PushPinNote2" | "Pumpkin1" | "EggsBlack" | "Cup" | "HeartGray" | "GingerbreadMan" | "BabyPacifier" | "BabyRattle" | "Cabins" | "HouseFunky" | "StarsBlack" | "Snowflakes" | "SnowflakeFancy" | "Skyrocket" | "Seattle" | "MusicNotes" | "PalmsBlack" | "MapleLeaf" | "PaperClips" | "ShorebirdTracks" | "People" | "PeopleWaving" | "EclipsingSquares2" | "Hypnotic" | "DiamondsGray" | "DecoArch" | "DecoBlocks" | "CirclesLines" | "Papyrus" | "Woodwork" | "WeavingBraid" | "WeavingRibbon" | "WeavingAngles" | "ArchedScallops" | "Safari" | "CelticKnotwork" | "CrazyMaze" | "EclipsingSquares1" | "Birds" | "FlowersTeacup" | "Northwest" | "Southwest" | "Tribal6" | "Tribal4" | "Tribal3" | "Tribal2" | "Tribal5" | "XIllusions" | "ZanyTriangles" | "Pyramids" | "PyramidsAbove" | "ConfettiGrays" | "ConfettiOutline" | "ConfettiWhite" | "Mosaic" | "Lightning2" | "HeebieJeebies" | "LightBulb" | "Gradient" | "TriangleParty" | "TwistedLines2" | "Moons" | "Ovals" | "DoubleDiamonds" | "ChainLink" | "Triangles" | "Tribal1" | "MarqueeToothed" | "SharksTeeth" | "Sawtooth" | "SawtoothGray" | "PostageStamp" | "WeavingStrips" | "ZigZag" | "CrossStitch" | "Gems" | "CirclesRectangles" | "CornerTriangles" | "CreaturesInsects" | "ZigZagStitch" | "Checkered" | "CheckedBarBlack" | "Marquee" | "BasicWhiteDots" | "BasicWideMidline" | "BasicWideOutline" | "BasicWideInline" | "BasicThinLines" | "BasicWhiteDashes" | "BasicWhiteSquares" | "BasicBlackSquares" | "BasicBlackDashes" | "BasicBlackDots" | "StarsTop" | "CertificateBanner" | "Handmade1" | "Handmade2" | "TornPaper" | "TornPaperBlack" | "CouponCutoutDashes" | "CouponCutoutDots"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### artWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the graphical page border specified in the artStyle property.

```typescript
artWidth?: number;
```

Property Value
number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### color

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
color?: string;
```

Property Value
string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### colorIndex

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal or [Word.Font](/en-us/javascript/api/word/word.font) object.

```typescript
colorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Property Value
[Word.ColorIndex](/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### inside

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if an inside border can be applied to the specified object.

```typescript
inside?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the border is visible.

```typescript
isVisible?: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line style of the border.

```typescript
lineStyle?: Word.BorderLineStyle | "None" | "Single" | "Dot" | "DashSmallGap" | "DashLargeGap" | "DashDot" | "DashDotDot" | "Double" | "Triple" | "ThinThickSmallGap" | "ThickThinSmallGap" | "ThinThickThinSmallGap" | "ThinThickMedGap" | "ThickThinMedGap" | "ThinThickThinMedGap" | "ThinThickLargeGap" | "ThickThinLargeGap" | "ThinThickThinLargeGap" | "SingleWavy" | "DoubleWavy" | "DashDotStroked" | "Emboss3D" | "Engrave3D" | "Outset" | "Inset";
```

Property Value
[Word.BorderLineStyle](/en-us/javascript/api/word/word.borderlinestyle) | "None" | "Single" | "Dot" | "DashSmallGap" | "DashLargeGap" | "DashDot" | "DashDotDot" | "Double" | "Triple" | "ThinThickSmallGap" | "ThickThinSmallGap" | "ThinThickThinSmallGap" | "ThinThickMedGap" | "ThickThinMedGap" | "ThinThickThinMedGap" | "ThinThickLargeGap" | "ThickThinLargeGap" | "ThinThickThinLargeGap" | "SingleWavy" | "DoubleWavy" | "DashDotStroked" | "Emboss3D" | "Engrave3D" | "Outset" | "Inset"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lineWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line width of an object's border.

```typescript
lineWidth?: Word.LineWidth | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600";
```

Property Value
[Word.LineWidth](/en-us/javascript/api/word/word.linewidth) | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)