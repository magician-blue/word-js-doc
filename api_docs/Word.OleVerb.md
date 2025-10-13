# Word.OleVerb enum

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the action associated with the verb that the OLE object should perform.

## Remarks

[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Fields

- discardUndoState = "DiscardUndoState"
  - Forces the object to discard any undo state that it might be maintaining; note that the object remains active, however.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- hide = "Hide"
  - Removes the object's user interface from view.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- inPlaceActivate = "InPlaceActivate"
  - Runs the object and installs its window, but doesn't install any user-interface tools.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- open = "Open"
  - Opens the object in a separate window.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- primary = "Primary"
  - Performs the verb that is invoked when the user double-clicks the object.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- show = "Show"
  - Shows the object to the user for editing or viewing. Use it to show a newly inserted object for initial editing.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- uiActivate = "UiActivate"
  - Activates the object in place and displays any user-interface tools that the object needs, such as menus or toolbars.
  - [ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]