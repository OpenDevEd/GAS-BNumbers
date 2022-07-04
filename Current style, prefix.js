function showCurrentStyle() {
  const ui = DocumentApp.getUi();
  const activeStyle = getHeadingStyle();
  Logger.log(activeStyle);
  onOpen();
  let menuItemText = headingStyles[activeStyle.value]['name'];
  //if (activeStyle.prefix == 'Custom') {
    menuItemText = menuItemText.replace('~PREFIX~', activeStyle.prefixText);
  //}

  ui.alert(menuItemText);
}


var headingStyles = {
  "numberHeadingsAddChapter23": {
    "name": "H1(~PREFIX~#)/H2-H3/-/figure",
    "separatorAbove": true,
    "replacePrefix": true,
    "xstyle" : ['1#', '1', '1', '-', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "OverrideH1PrefixWithCustomPrefix": true,
    "numberDepth": 3
  },
  "numberHeadingsAddChapter234": {
    "name": "H1(~PREFIX~#)/H2-H4/-/figure",
    "replacePrefix": true,
    "subMenuBelow": true,
    "xstyle" : ['1#', '1', '1', '1', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "OverrideH1PrefixWithCustomPrefix": true, 
    "numberDepth": 4
  },
  "numberHeadingsAdd3Figure": {
    "name": "H1-3/-/-/figure",
    "separatorAbove": true,
    "xstyle" : ['1', '1', '1', '-', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 3
  },
  "numberHeadingsAdd4Figure": {
    "name": "H1-4/-/figure",
    "xstyle" : ['1', '1', '1', '1', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 4
    },
  "numberHeadingsAddFigure": {
    "name": "H1-5/figure",
    "xstyle" : ['1', '1', '1', '1', '1', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 5
  },
  "numberHeadingsAdd1WithLinksOrNot": {
    "name": "H1",
    "separatorAbove": true,
    "xstyle" : ['1', '-', '-', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 1
  },
  "numberHeadingsAdd2WithLinksOrNot": {
    "name": "H1-2",
    "xstyle" : ['1', '1', '-', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 2
  },
  "numberHeadingsAdd3WithLinksOrNot": {
    "name": "H1-3",
    "xstyle" : ['1', '1', '1', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 3
  },
  "numberHeadingsAddPartial4WithLinksOrNot": {
    "name": "H1-4",
    "xstyle" : ['1', '1', '1', '1', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 4
  },
  "numberHeadingsAdd5WithLinksOrNot": {
    "name": "H1-5",
    "xstyle" : ['1', '1', '1', '1', '1', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 5
  },
  "numberHeadingsAdd6WithLinksOrNot": {
    "name": "H1-6",
    "xstyle" : ['1', '1', '1', '1', '1', '1'],
    "prefixlead": [null, null, null, null, null, null],
    "OverrideH1PrefixWithCustomPrefix": false, 
    "numberDepth": 6
  }
};


var headingStylesOLD = {
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


