# Word.Interfaces.TabStopData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `tabStop.toJSON()`.

## Properties

| Property   | Description                                                                 |
|------------|-----------------------------------------------------------------------------|
| alignment  | Gets a `TabAlignment` value that represents the alignment for the tab stop. |
| customTab  | Gets whether this tab stop is a custom tab stop.                            |
| leader     | Gets a `TabLeader` value that represents the leader for this `TabStop` object. |
| position   | Gets the position of the tab stop relative to the left margin.              |

## Property Details

### alignment

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TabAlignment` value that represents the alignment for the tab stop.

```typescript
alignment?: Word.TabAlignment | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List";
```

#### Property Value
[Word.TabAlignment](/en-us/javascript/api/word/word.tabalignment) | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### customTab

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether this tab stop is a custom tab stop.

```typescript
customTab?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### leader

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TabLeader` value that represents the leader for this `TabStop` object.

```typescript
leader?: Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot";
```

#### Property Value
[Word.TabLeader](/en-us/javascript/api/word/word.tableader) | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### position

Note

This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the tab stop relative to the left margin.

```typescript
position?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)