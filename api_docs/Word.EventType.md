# Word.EventType enum

Package: [word](/en-us/javascript/api/word)

Provides information about the type of a raised event.

## Remarks

[ API set: WordApi 1.5 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml

async function contentControlDeleted(event: Word.ContentControlDeletedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were deleted:`, event.ids);
  });
}
```

## Fields

- annotationClicked = "AnnotationClicked"
  - Represents that an annotation was clicked (or selected with **Alt+Down**) in the document.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- annotationHovered = "AnnotationHovered"
  - Represents that an annotation was hovered over in the document.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- annotationInserted = "AnnotationInserted"
  - Represents that one or more annotations were added in the document.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- annotationPopupAction = "AnnotationPopupAction"
  - Represents an action in the annotation pop-up.
  - [ API set: WordApi 1.8 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- annotationRemoved = "AnnotationRemoved"
  - Represents that one or more annotations were deleted from the document.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- commentAdded = "CommentAdded"
  - Represents that one or more new comments were added.
  - [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- commentChanged = "CommentChanged"
  - Represents that a comment or its reply was changed.
  - [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- commentDeleted = "CommentDeleted"
  - Represents that one or more comments were deleted.
  - [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- commentDeselected = "CommentDeselected"
  - Represents that a comment was deselected.
  - [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- commentSelected = "CommentSelected"
  - Represents that a comment was selected.
  - [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlAdded = "ContentControlAdded"
  - ContentControlAdded represents the event a content control has been added to the document.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlDataChanged = "ContentControlDataChanged"
  - ContentControlDataChanged represents the event that the data in the content control have been changed.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlDeleted = "ContentControlDeleted"
  - ContentControlDeleted represents the event that the content control has been deleted.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlEntered = "ContentControlEntered"
  - Represents that a content control has been entered.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlExited = "ContentControlExited"
  - Represents that a content control has been exited.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- contentControlSelectionChanged = "ContentControlSelectionChanged"
  - ContentControlSelectionChanged represents the event that the selection in the content control has been changed.
  - [ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- paragraphAdded = "ParagraphAdded"
  - Represents that one or more new paragraphs were added.
  - [ API set: WordApi 1.6 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- paragraphChanged = "ParagraphChanged"
  - Represents that one or more paragraphs were changed.
  - [ API set: WordApi 1.6 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- paragraphDeleted = "ParagraphDeleted"
  - Represents that one or more paragraphs were deleted.
  - [ API set: WordApi 1.6 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)