function setPurpleForegroundNotes() {
  setNotesColours(null, '#9900ff');
}
function setYellowBackgroundNotes() {
  setNotesColours('#ffff00', null);
}

function setNotesColours(backgroundColor, foregroundColor) {
  var bodyElement = DocumentApp.getActiveDocument().getBody();
  const { status, dividerEl } = notesExistenceCheck(bodyElement);
  if (!status) return 0;

  var NOTE_RE_DOC = "^[ \t\n\r\f\v]*⟦[0-9]+⟧[ \t\n\r\f\v]*.*$";
  var searchResult = bodyElement.findText(NOTE_RE_DOC, dividerEl);

  if (searchResult == null) {
    return 0;
  }

  const elToColour = [];
  var thisElement = searchResult.getElement();
  const childIndex = bodyElement.getChildIndex(thisElement.getParent());

  elToColour.push(thisElement);
  const numChildren = bodyElement.getNumChildren();

  for (let i = childIndex + 1; i < numChildren; i++) {
    const currentEl = bodyElement.getChild(i);
    if (currentEl.getType() === DocumentApp.ElementType.PARAGRAPH) {
      elToColour.push(currentEl);
    }
  }

  elToColour.forEach(thisElement => {
    var thisElementText = thisElement.asText();
    thisElementText.setBackgroundColor(backgroundColor);
    thisElementText.setForegroundColor(foregroundColor);
  });

  /*
  while (searchResult !== null) {
    var thisElement = searchResult.getElement();
    var thisElementText = thisElement.asText();

    if (backgroundColor) {
      thisElementText.setBackgroundColor(backgroundColor);
    };
    if (foregroundColor) {
      thisElementText.setForegroundColor(foregroundColor);
    };
    // search for next match
    searchResult = bodyElement.findText(NOTE_RE_DOC, searchResult);
  }
  */

}