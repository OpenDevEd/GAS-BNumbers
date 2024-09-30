function clearBoldFigureH6Settings() {
  const docProperties = PropertiesService.getDocumentProperties();
  docProperties.deleteProperty('BOLD_FIGURE_H6_SETTINGS');
  onOpen();
}
function testBoldFigureH6Settings() {
  Logger.log(getDocumentPropertyString('BOLD_FIGURE_H6_SETTINGS'));
}

function activateBoldFigureH6Settings() {
  activateSettingsYesNo('BOLD_FIGURE_H6_SETTINGS');
}

function activateSettingsYesNo(propertyKey) {
  const savedSettings = getDocumentPropertyString(propertyKey);
  const value = savedSettings === 'yes' || savedSettings == null ? 'no' : 'yes';
  setDocumentPropertyString(propertyKey, value);
  if (value === 'yes'){
    reformatHeadings6();
  }
  onOpen();
}

function getBoldFigureStyle(tryToRetrieveProperties) {
  const resultObj = {
    marker: '',
    style: 'yes',
    menuText: headingStyles['boldFugureH6']['name']
  };
  if (tryToRetrieveProperties === true) {
    try {
      const savedSettings = getDocumentPropertyString('BOLD_FIGURE_H6_SETTINGS');
      if (savedSettings == null) {
        resultObj['marker'] = '◯';
      } else if (savedSettings === 'yes'){
        resultObj['marker'] = '🟢';
      }else{
        resultObj['style'] = savedSettings;
      }
    }
    catch (error) {
      // Logger.log('Needs to activate!!! ' + error);
    }
  }
  // Logger.log(resultObj);
  return resultObj;
}

function reformatHeadings6() {
  const doc = DocumentApp.getActiveDocument();
  const p = doc.getParagraphs();
  const style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.ITALIC] = false;

  for (let i in p) {
    const e = p[i];
    let eText = e.getText() + '';
    const eTypeString = e.getHeading().toString();

    if (!eTypeString.match(/Heading ?6/i)) {
      // continue if the paragraph is not a heading
      continue;
    }

    eText = eText.trim();

    if (eText == '') {
      // continue if the paragraph is empty
      continue;
    }
    // Logger.log(eText);

    const checkTableBoxFigure = /^(Figure|Box|Table) (\d+|X)\.?\d*\. /.exec(eText);
    if (checkTableBoxFigure != null) {
      // Logger.log(checkTableBoxFigure);
      e.editAsText().setAttributes(0, checkTableBoxFigure[0].length - 1, style)
    }
  }
}
