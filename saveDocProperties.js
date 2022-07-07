// savePrefixToDocumentProperty, activateHeadingStyle use the function
function setDocumentPropertyString(property_name, value) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(property_name, value);
}

// getHeadingStyle uses the function
function getDocumentPropertyString(property_name) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var value = documentProperties.getProperty(property_name);
  return value;
}
