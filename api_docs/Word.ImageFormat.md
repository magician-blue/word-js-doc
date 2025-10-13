# Word.ImageFormat enum

Package: [word](/en-us/javascript/api/word)

## Remarks

[ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Gets the first image in the document.
await Word.run(async (context) => {
  const firstPicture: Word.InlinePicture = context.document.body.inlinePictures.getFirst();
  firstPicture.load("width, height, imageFormat");

  await context.sync();
  console.log(`Image dimensions: ${firstPicture.width} x ${firstPicture.height}`, `Image format: ${firstPicture.imageFormat}`);
  // Get the image encoded as Base64.
  const base64 = firstPicture.getBase64ImageSrc();

  await context.sync();
  console.log(base64.value);
});
```

## Fields

- bmp = "Bmp"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- emf = "Emf"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- exif = "Exif"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- gif = "Gif"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- icon = "Icon"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- jpeg = "Jpeg"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pdf = "Pdf"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pict = "Pict"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- png = "Png"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- svg = "Svg"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- tiff = "Tiff"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- undefined = "Undefined"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- unsupported = "Unsupported"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- wmf = "Wmf"
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)