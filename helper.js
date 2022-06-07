
function getValueFromUser(title,text, defaultOK, defaultCancel, defaultClose) {
  text = text || "Please enter a value.";
  defaultOK = defaultOK || "";
  defaultCancel = defaultCancel || null;
  defaultClose = defaultClose || null;
  if (!text) {
    //text = text || title;
    title = text;
    title = "BUtils";
  };
  var result = DocumentApp.getUi().prompt(title,text, DocumentApp.getUi().ButtonSet.OK_CANCEL);
  // Process the user's response:
  if (result.getSelectedButton() == DocumentApp.getUi().Button.OK) {
    var res = result.getResponseText();
    // DocumentApp.getUi().alert('Result: '+res);
    if (res == "" && defaultOK) {
      return defaultOK;
    } else {      
      return res;
    };
  } else if (result.getSelectedButton() == DocumentApp.getUi().Button.CANCEL) {
    //DocumentApp.getUi().alert('The user didn\'t want to provide a value.');
    return defaultCancel;
  } else if (result.getSelectedButton() == DocumentApp.getUi().Button.CLOSE) {
    //DocumentApp.getUi().alert('The user clicked the close button in the dialog\'s title bar.');
    return defaultClose;
  }
  DocumentApp.getUi().alert('Unknown action.');
  return null;
}

function getConfirmationFromUser(text) {
  // Display a dialog box with a message and "Yes" and "No" buttons.
  var ui = DocumentApp.getUi();
  var response = ui.alert(text, ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  if (response == ui.Button.OK) {
    return true;
  } else {
    return false;
  }
}


function alert(text) {
  DocumentApp.getUi().alert(text);
};

function alertv(text,myvariable) {
  DocumentApp.getUi().alert(text+" = "+JSON.stringify(myvariable));
};

function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index;
};


