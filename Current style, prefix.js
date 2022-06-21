function showCurrentStyle() {
  const ui = DocumentApp.getUi();
  const activeStyle = getHeadingStyle();
  onOpen();
  let menuItemText = headingStyles[activeStyle.value]['name'];
  if (activeStyle.prefix == 'Custom') {
    menuItemText = menuItemText.replace('~PREFIX~', activeStyle.prefixText);
  }

  ui.alert(menuItemText);
}


var headingStyles = {
  "numberHeadingsAdd3WithLinks": {
    "name": "H1-3 (update links)",
    "separatorAbove": true,
  },
  "numberHeadingsAdd3Figure": {
    "name": "H1-3/-/-/figure (update links)"
  },
  "numberHeadingsAdd4Figure": {
    "name": "H1-4/-/figure (update links)",
  },
  "numberHeadingsAddFigure": {
    "name": "H1-5/figure (update links)"
  },
  "numberHeadingsAddChapter23": {
    "name": "H1(~PREFIX~#)/H2-H3/-/figure (update links)",
    "separatorAbove": true,
    "replacePrefix": true
  },
  "numberHeadingsAddChapter234": {
    "name": "H1(~PREFIX~#)/H2-H4/-/figure (update links)",
    "subMenuBelow": true,
    "replacePrefix": true
  },
  "numberHeadingsAddPartial3": {
    "name": "H1-3 (keep links)",
    "separatorAbove": true
  },
  "numberHeadingsAddPartial4": {
    "name": "H1-4 (keep links)"
  },
  "numberHeadingsAdd6": {
    "name": "H1-6 (keep links)"
  }
};


var prefixes = {
  "None": {
    name: "<none>",
    func: "setPrefixToNone",
    value: ""
  },
  "Chapter": {
    "name": "Chapter_",
    func: "setPrefixToChapter",
    value: "Chapter "
  },
  "Section": {
    "name": "Section_",
    func: "setPrefixToSection",
    value: "Section "
  },
  "Session": {
    "name": "Session_",
    func: "setPrefixToSession",
    value: "Session "
  },
  "Custom": {
    "name": "<enter>",
    func: "enterPrefix"
  }
};

function setPrefixToNone() {
  savePrefixToDocumentProperty("None");
}

function setPrefixToChapter() {
  savePrefixToDocumentProperty("Chapter");
}
function setPrefixToSection() {
  savePrefixToDocumentProperty("Section");
}
function setPrefixToSession() {
  savePrefixToDocumentProperty("Session");
}
function enterPrefix() {
  //savePrefixToDocumentProperty("None");

  var prefix = getValueFromUser('User defined prefix', 'Please type space after the string, if you want a space before the number.');
  if (prefix != null & prefix != '') {
    Logger.log('prefix=' + prefix);
    setDocumentPropertyString('BNumbers_Custom_Prefix_Property', prefix);
    savePrefixToDocumentProperty('Custom');
  } else {
    Logger.log('?');
  }

}


