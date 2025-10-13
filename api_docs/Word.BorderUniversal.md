# Word.BorderUniversal class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the BorderUniversal object, which manages borders for a range, paragraph, table, or frame.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]
https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Properties
- artStyle: Specifies the graphical page-border design for the document.
- artWidth: Specifies the width (in points) of the graphical page border specified in the artStyle property.
- color: Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.
- colorIndex: Specifies the color for the BorderUniversal or Word.Font object.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- inside: Returns true if an inside border can be applied to the specified object.
- isVisible: Specifies whether the border is visible.
- lineStyle: Specifies the line style of the border.
- lineWidth: Specifies the line width of an object's border.

## Methods
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderUniversalData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### artStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the graphical page-border design for the document.

```typescript
artStyle: Word.PageBorderArt | "Apples" | "MapleMuffins" | "CakeSlice" | "CandyCorn" | "IceCreamCones" | "ChampagneBottle" | "PartyGlass" | "ChristmasTree" | "Trees" | "PalmsColor" | "Balloons3Colors" | "BalloonsHotAir" | "PartyFavor" | "ConfettiStreamers" | "Hearts" | "HeartBalloon" | "Stars3D" | "StarsShadowed" | "Stars" | "Sun" | "Earth2" | "Earth1" | "PeopleHats" | "Sombrero" | "Pencils" | "Packages" | "Clocks" | "Firecrackers" | "Rings" | "MapPins" | "Confetti" | "CreaturesButterfly" | "CreaturesLadyBug" | "CreaturesFish" | "BirdsFlight" | "ScaredCat" | "Bats" | "FlowersRoses" | "FlowersRedRose" | "Poinsettias" | "Holly" | "FlowersTiny" | "FlowersPansy" | "FlowersModern2" | "FlowersModern1" | "WhiteFlowers" | "Vine" | "FlowersDaisies" | "FlowersBlockPrint" | "DecoArchColor" | "Fans" | "Film" | "Lightning1" | "Compass" | "DoubleD" | "ClassicalWave" | "ShadowedSquares" | "TwistedLines1" | "Waveline" | "Quadrants" | "CheckedBarColor" | "Swirligig" | "PushPinNote1" | "PushPinNote2" | "Pumpkin1" | "EggsBlack" | "Cup" | "HeartGray" | "GingerbreadMan" | "BabyPacifier" | "BabyRattle" | "Cabins" | "HouseFunky" | "StarsBlack" | "Snowflakes" | "SnowflakeFancy" | "Skyrocket" | "Seattle" | "MusicNotes" | "PalmsBlack" | "MapleLeaf" | "PaperClips" | "ShorebirdTracks" | "People" | "PeopleWaving" | "EclipsingSquares2" | "Hypnotic" | "DiamondsGray" | "DecoArch" | "DecoBlocks" | "CirclesLines" | "Papyrus" | "Woodwork" | "WeavingBraid" | "WeavingRibbon" | "WeavingAngles" | "ArchedScallops" | "Safari" | "CelticKnotwork" | "CrazyMaze" | "EclipsingSquares1" | "Birds" | "FlowersTeacup" | "Northwest" | "Southwest" | "Tribal6" | "Tribal4" | "Tribal3" | "Tribal2" | "Tribal5" | "XIllusions" | "ZanyTriangles" | "Pyramids" | "PyramidsAbove" | "ConfettiGrays" | "ConfettiOutline" | "ConfettiWhite" | "Mosaic" | "Lightning2" | "HeebieJeebies" | "LightBulb" | "Gradient" | "TriangleParty" | "TwistedLines2" | "Moons" | "Ovals" | "DoubleDiamonds" | "ChainLink" | "Triangles" | "Tribal1" | "MarqueeToothed" | "SharksTeeth" | "Sawtooth" | "SawtoothGray" | "PostageStamp" | "WeavingStrips" | "ZigZag" | "CrossStitch" | "Gems" | "CirclesRectangles" | "CornerTriangles" | "CreaturesInsects" | "ZigZagStitch" | "Checkered" | "CheckedBarBlack" | "Marquee" | "BasicWhiteDots" | "BasicWideMidline" | "BasicWideOutline" | "BasicWideInline" | "BasicThinLines" | "BasicWhiteDashes" | "BasicWhiteSquares" | "BasicBlackSquares" | "BasicBlackDashes" | "BasicBlackDots" | "StarsTop" | "CertificateBanner" | "Handmade1" | "Handmade2" | "TornPaper" | "TornPaperBlack" | "CouponCutoutDashes" | "CouponCutoutDots";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.pageborderart | "Apples" | "MapleMuffins" | "CakeSlice" | "CandyCorn" | "IceCreamCones" | "ChampagneBottle" | "PartyGlass" | "ChristmasTree" | "Trees" | "PalmsColor" | "Balloons3Colors" | "BalloonsHotAir" | "PartyFavor" | "ConfettiStreamers" | "Hearts" | "HeartBalloon" | "Stars3D" | "StarsShadowed" | "Stars" | "Sun" | "Earth2" | "Earth1" | "PeopleHats" | "Sombrero" | "Pencils" | "Packages" | "Clocks" | "Firecrackers" | "Rings" | "MapPins" | "Confetti" | "CreaturesButterfly" | "CreaturesLadyBug" | "CreaturesFish" | "BirdsFlight" | "ScaredCat" | "Bats" | "FlowersRoses" | "FlowersRedRose" | "Poinsettias" | "Holly" | "FlowersTiny" | "FlowersPansy" | "FlowersModern2" | "FlowersModern1" | "WhiteFlowers" | "Vine" | "FlowersDaisies" | "FlowersBlockPrint" | "DecoArchColor" | "Fans" | "Film" | "Lightning1" | "Compass" | "DoubleD" | "ClassicalWave" | "ShadowedSquares" | "TwistedLines1" | "Waveline" | "Quadrants" | "CheckedBarColor" | "Swirligig" | "PushPinNote1" | "PushPinNote2" | "Pumpkin1" | "EggsBlack" | "Cup" | "HeartGray" | "GingerbreadMan" | "BabyPacifier" | "BabyRattle" | "Cabins" | "HouseFunky" | "StarsBlack" | "Snowflakes" | "SnowflakeFancy" | "Skyrocket" | "Seattle" | "MusicNotes" | "PalmsBlack" | "MapleLeaf" | "PaperClips" | "ShorebirdTracks" | "People" | "PeopleWaving" | "EclipsingSquares2" | "Hypnotic" | "DiamondsGray" | "DecoArch" | "DecoBlocks" | "CirclesLines" | "Papyrus" | "Woodwork" | "WeavingBraid" | "WeavingRibbon" | "WeavingAngles" | "ArchedScallops" | "Safari" | "CelticKnotwork" | "CrazyMaze" | "EclipsingSquares1" | "Birds" | "FlowersTeacup" | "Northwest" | "Southwest" | "Tribal6" | "Tribal4" | "Tribal3" | "Tribal2" | "Tribal5" | "XIllusions" | "ZanyTriangles" | "Pyramids" | "PyramidsAbove" | "ConfettiGrays" | "ConfettiOutline" | "ConfettiWhite" | "Mosaic" | "Lightning2" | "HeebieJeebies" | "LightBulb" | "Gradient" | "TriangleParty" | "TwistedLines2" | "Moons" | "Ovals" | "DoubleDiamonds" | "ChainLink" | "Triangles" | "Tribal1" | "MarqueeToothed" | "SharksTeeth" | "Sawtooth" | "SawtoothGray" | "PostageStamp" | "WeavingStrips" | "ZigZag" | "CrossStitch" | "Gems" | "CirclesRectangles" | "CornerTriangles" | "CreaturesInsects" | "ZigZagStitch" | "Checkered" | "CheckedBarBlack" | "Marquee" | "BasicWhiteDots" | "BasicWideMidline" | "BasicWideOutline" | "BasicWideInline" | "BasicThinLines" | "BasicWhiteDashes" | "BasicWhiteSquares" | "BasicBlackSquares" | "BasicBlackDashes" | "BasicBlackDots" | "StarsTop" | "CertificateBanner" | "Handmade1" | "Handmade2" | "TornPaper" | "TornPaperBlack" | "CouponCutoutDashes" | "CouponCutoutDots"

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### artWidth
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the graphical page border specified in the artStyle property.

```typescript
artWidth: number;
```

Type: number

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### color
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
color: string;
```

Type: string

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### colorIndex
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the BorderUniversal or Word.Font object.

```typescript
colorIndex: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### inside
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if an inside border can be applied to the specified object.

```typescript
readonly inside: boolean;
```

Type: boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isVisible
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the border is visible.

```typescript
isVisible: boolean;
```

Type: boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### lineStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line style of the border.

```typescript
lineStyle: Word.BorderLineStyle | "None" | "Single" | "Dot" | "DashSmallGap" | "DashLargeGap" | "DashDot" | "DashDotDot" | "Double" | "Triple" | "ThinThickSmallGap" | "ThickThinSmallGap" | "ThinThickThinSmallGap" | "ThinThickMedGap" | "ThickThinMedGap" | "ThinThickThinMedGap" | "ThinThickLargeGap" | "ThickThinLargeGap" | "ThinThickThinLargeGap" | "SingleWavy" | "DoubleWavy" | "DashDotStroked" | "Emboss3D" | "Engrave3D" | "Outset" | "Inset";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.borderlinestyle | "None" | "Single" | "Dot" | "DashSmallGap" | "DashLargeGap" | "DashDot" | "DashDotDot" | "Double" | "Triple" | "ThinThickSmallGap" | "ThickThinSmallGap" | "ThinThickThinSmallGap" | "ThinThickMedGap" | "ThickThinMedGap" | "ThinThickThinMedGap" | "ThinThickLargeGap" | "ThickThinLargeGap" | "ThinThickThinLargeGap" | "SingleWavy" | "DoubleWavy" | "DashDotStroked" | "Emboss3D" | "Engrave3D" | "Outset" | "Inset"

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### lineWidth
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line width of an object's border.

```typescript
lineWidth: Word.LineWidth | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600";
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.linewidth | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600"

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Method Details

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BorderUniversalLoadOptions): Word.BorderUniversal;
```

Parameters
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversalloadoptions  
  Provides options for which properties of the object to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BorderUniversal;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.BorderUniversal;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal

### set(properties, options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.BorderUniversalUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversalupdatedata  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.BorderUniversal): void;
```

Parameters
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal

Returns
- void

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderUniversalData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BorderUniversalData;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversaldata

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BorderUniversal;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BorderUniversal;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversal