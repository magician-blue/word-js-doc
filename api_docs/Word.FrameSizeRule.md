# Word.FrameSizeRule enum

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents how Word interprets the rule used to determine the height or width of a [Word.Frame](/en-us/javascript/api/word/word.frame).

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Fields

- atLeast = "AtLeast"
  - The height or width is set to a value equal to or greater than the value specified by the `height` property or `width` property.
  - [ API set: WordApi BETA (PREVIEW ONLY) ]

- auto = "Auto"
  - The height or width is set according to the height or width of the item in the frame.
  - [ API set: WordApi BETA (PREVIEW ONLY) ]

- exact = "Exact"
  - The height or width is set to an exact value specified by the `height` property or `width` property.
  - [ API set: WordApi BETA (PREVIEW ONLY) ]