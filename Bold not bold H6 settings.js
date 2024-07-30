function clearBoldFigureH6Settings() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('BOLD_FIGURE_H6_SETTINGS');
  onOpen();
}
function testBoldFigureH6Settings() {
  Logger.log(PropertiesService.getUserProperties().getProperty('BOLD_FIGURE_H6_SETTINGS'));
}

function activateBoldFigureH6Settings() {
  activateSettingsYesNo('BOLD_FIGURE_H6_SETTINGS');
}

function activateSettingsYesNo(propertyKey) {
  // const value = Object.keys(allMenuItemsObj).find(key => allMenuItemsObj[key] === targetObj);
  // const currentValue = getSettings(tryToRetrieveProperties = true, defaultSettings = 'yes', allMenuItemsObj, propertyKey)

  const savedSettings = PropertiesService.getUserProperties().getProperty(propertyKey);
  const value = savedSettings === 'yes' || savedSettings == null ? 'no' : 'yes';
  PropertiesService.getUserProperties().setProperty(propertyKey, value);
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
      const savedSettings = PropertiesService.getUserProperties().getProperty('BOLD_FIGURE_H6_SETTINGS');
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