# Word.Interfaces.LineNumberingUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the LineNumbering object, for use in lineNumbering.set({ ... }).

## Properties

- [countBy](#word-word-interfaces-linenumberingupdatedata-countby-member) — Specifies the numeric increment for line numbers.
- [distanceFromText](#word-word-interfaces-linenumberingupdatedata-distancefromtext-member) — Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.
- [isActive](#word-word-interfaces-linenumberingupdatedata-isactive-member) — Specifies if line numbering is active for the specified document, section, or sections.
- [restartMode](#word-word-interfaces-linenumberingupdatedata-restartmode-member) — Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.
- [startingNumber](#word-word-interfaces-linenumberingupdatedata-startingnumber-member) — Specifies the starting line number.

## Property Details

<a id="word-word-interfaces-linenumberingupdatedata-countby-member"></a>
### countBy

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the numeric increment for line numbers.

```typescript
countBy?: number;
```

#### Property Value
number

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-linenumberingupdatedata-distancefromtext-member"></a>
### distanceFromText

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.

```typescript
distanceFromText?: number;
```

#### Property Value
number

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-linenumberingupdatedata-isactive-member"></a>
### isActive

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if line numbering is active for the specified document, section, or sections.

```typescript
isActive?: boolean;
```

#### Property Value
boolean

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-linenumberingupdatedata-restartmode-member"></a>
### restartMode

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.

```typescript
restartMode?: Word.NumberingRule | "RestartContinuous" | "RestartSection" | "RestartPage";
```

#### Property Value
[Word.NumberingRule](/en-us/javascript/api/word/word.numberingrule) | "RestartContinuous" | "RestartSection" | "RestartPage"

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-linenumberingupdatedata-startingnumber-member"></a>
### startingNumber

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting line number.

```typescript
startingNumber?: number;
```

#### Property Value
number

#### Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]