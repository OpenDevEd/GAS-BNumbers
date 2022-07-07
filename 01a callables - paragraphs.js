// Menu item Paragraphs -> 'pna Paragraph numbers add, with  ⟦ and ⟧'
// Numbers paragraphs
function paraNumAdd() {
  sequentialNumbers(true, false, 0);
}

// Menu item Paragraphs -> 'psna Paragraph/sentence numbers add, with ⟦ and ⟧'
// Numbers paragraphs and sentences.
function paraSenNumAdd() {
  sequentialNumbersX(true, true, 2, 0);
}

// Menu item Paragraphs -> 'psnr Paragraph/sentence numbers remove'
// Removes paragraph/sentence numbers.
function paraSenNumRemove() {
  sequentialNumbers(false, true, 0);
}

// Menu item Paragraphs -> 'psnm Paragraph/sentence numbers minify'
// Applies font size 1pt, white colour, and empty background to paragraph/sentence numbers.
function minifyParaSenMarker() {
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FEFEFE'; // null
  //style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#AAAAAA'; // null
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  style[DocumentApp.Attribute.FONT_SIZE] = 1;
  regexpRestyle("⟦[^⟦⟧]+⟧", style);
}

// Menu item Paragraphs -> 'psnshow Paragraph/sentence numbers - restore size'
// Applies font size 10pt, black colour, and yellow background to paragraph/sentence numbers.
function maxifyParaSenMarker() {
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#111111'; // null
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = "#FFFF00";
  style[DocumentApp.Attribute.FONT_SIZE] = 10;
  regexpRestyle("⟦[^⟦⟧]+⟧", style);
}