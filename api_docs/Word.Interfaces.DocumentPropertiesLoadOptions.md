# Word.Interfaces.DocumentPropertiesLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents document properties.

## Remarks

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- applicationName — Gets the application name of the document.
- author — Specifies the author of the document.
- category — Specifies the category of the document.
- comments — Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.
- company — Specifies the company of the document.
- creationDate — Gets the creation date of the document.
- format — Specifies the format of the document.
- keywords — Specifies the keywords of the document.
- lastAuthor — Gets the last author of the document.
- lastPrintDate — Gets the last print date of the document.
- lastSaveTime — Gets the last save time of the document.
- manager — Specifies the manager of the document.
- revisionNumber — Gets the revision number of the document.
- security — Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #