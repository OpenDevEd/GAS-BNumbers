function clearWithWithoutDotsSettings() {
  const docProperties = PropertiesService.getDocumentProperties();
  docProperties.deleteProperty('WITH_WITHOUT_DOTS_SETTINGS');
  onOpen();
}

function testWithWithoutDotsSettings() {
  Logger.log(PropertiesService.getDocumentProperties().getProperty('WITH_WITHOUT_DOTS_SETTINGS'));
}

function activateWithWithoutDotsSettings(obj) {
  activateSettings(withWithoutDotsStyles, obj, 'WITH_WITHOUT_DOTS_SETTINGS');
  updateNumbersInLinksToHeadings();
}

// Menu items of with/without dots settings
const withWithoutDotsStyles = {
  "withDots": {
    "menuText": "Heading numbering with dots in links",
    "run": function () { activateWithWithoutDotsSettings(this); }
  },
  "withoutDots": {
    "menuText": "Heading numbering without dots in links",
    "run": function () { activateWithWithoutDotsSettings(this); }
  }
}
