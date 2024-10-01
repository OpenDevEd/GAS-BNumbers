function greenCheckboxMenuHelper(TrMenu, tryToRetrieveProperties, defaultSettings, allMenuItemsObj, allMenuItemsObjName, propertyKey) {
  const activeSettings = getSettings(tryToRetrieveProperties, defaultSettings, allMenuItemsObj, propertyKey);
  let selectedSettingsMarker = '';
  for (let key in allMenuItemsObj) {
    selectedSettingsMarker = key === activeSettings.style ? ' ' + activeSettings.marker: '';
    TrMenu.addItem(allMenuItemsObj[key].menuText + selectedSettingsMarker, allMenuItemsObjName + '.' + key + '.run');
  }
}

function getSettings(tryToRetrieveProperties, defaultSettings, allMenuItemsObj, propertyKey) {
  const resultObj = {
    marker: '◯',
    style: defaultSettings,
    menuText: allMenuItemsObj[defaultSettings]['menuText']
  };
  if (tryToRetrieveProperties === true) {
    try {
      const savedSettings = getDocumentPropertyString(propertyKey);
      // const savedSettings = PropertiesService.getUserProperties().getProperty(propertyKey);
      if (savedSettings != null && allMenuItemsObj.hasOwnProperty(savedSettings)) {
        resultObj['style'] = savedSettings;
        resultObj['menuText'] = allMenuItemsObj[savedSettings]['menuText'];
        resultObj['marker'] = '🟢';
      }
    }
    catch (error) {
      // Logger.log('Needs to activate!!! ' + error);
    }
  }
  return resultObj;
}

function activateSettings(allMenuItemsObj, targetObj, propertyKey) {
  const value = Object.keys(allMenuItemsObj).find(key => allMenuItemsObj[key] === targetObj);

  // const userProperties = PropertiesService.getUserProperties();
  // userProperties.setProperty(propertyKey, value);
  setDocumentPropertyString(propertyKey, value);
  onOpen();
}
