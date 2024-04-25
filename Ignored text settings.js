function ignoredTextShowCurrentSetting() {
  const IGNORED_TEXT = getIgnoredText();
  let text = '';
  if (IGNORED_TEXT === '') {
    text = 'No headings will be ignored. \n\nDo you want to set regular expression to exclude headings from numbering?'
  } else {
    text = 'Headings matching this expression will not be numbered:\n\n' + IGNORED_TEXT + '\n\nDo you want to change the expression?'
  }

  const userAnswer = getConfirmationFromUser(text);
  if (userAnswer === true) {
    ignoredTextEditSetting(IGNORED_TEXT);
  }
}

function getIgnoredText() {
  let IGNORED_TEXT = getDocumentPropertyString('IGNORED_TEXT');
  if (IGNORED_TEXT == null) {
    IGNORED_TEXT = DEFAULT_IGNORED_TEXT;
  }
  return IGNORED_TEXT;
}

function ignoredTextEditSetting(currentRegEx, error = null) {
  if (currentRegEx == null) {
    currentRegEx = getIgnoredText();
    if (currentRegEx === '') {
      currentRegEx = 'No headings will be ignored.';
    }
  }
  const errorText = error == null ? '' : 'ERROR: ' + error + '\n\n';
  let newIgnoredText = getValueFromUser('Edit ignored text', errorText + 'Current expression: ' + currentRegEx + '\n\nPlease enter a regular expression.\nHeadings matching this expression will not be numbered.\n Leave the box empty if you do not want any headings to be ignored.', '');
  // Logger.log(newIgnoredText);
  if (newIgnoredText == null) {
    return 0;
  }
  newIgnoredText = newIgnoredText.trim();
  if (/^(\/\^.*\$\/[gimyu]*|)$/.test(newIgnoredText)) {
    saveIgnoredTextRegex(newIgnoredText);
  } else {
    ignoredTextEditSetting(currentRegEx, 'Enter regular expression in the format /^(text1|text2|text3)$/i or leave the box empty.\nYou entered: ' + newIgnoredText);
  }
}

function ignoredTextRestoreDefault() {
  PropertiesService.getDocumentProperties().deleteProperty('IGNORED_TEXT');
  alert('Regular expression was set to default:\n\n' + DEFAULT_IGNORED_TEXT);
}

function saveIgnoredTextRegex(newRegEx = '') {
  setDocumentPropertyString('IGNORED_TEXT', newRegEx);
}

function getIgnoredTextRegex() {
  const IGNORED_TEXT = getIgnoredText();
  if (IGNORED_TEXT === '') {
    return '';
  }
  // Extract the pattern and flags from the string
  const parts = IGNORED_TEXT.split('/');
  const pattern = parts[1];
  const flags = parts[2] || '';

  // Create the regular expression object
  const ignoredTextRegex = new RegExp(pattern, flags);
  return ignoredTextRegex;
}