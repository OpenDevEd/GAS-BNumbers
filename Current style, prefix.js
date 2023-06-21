// Menu item Select style -> 'Show current style'
// Shows the current (selected by user or default) style in alert
function showCurrentStyle() {
  //const ui = DocumentApp.getUi();
  const activeStyle = getHeadingStyle();
  onOpen();
  let menuItemText = headingStyles[activeStyle.value]['name'];
  menuItemText = menuItemText.replaceAll('~PREFIX~', activeStyle.prefixText);
  //ui.alert(menuItemText);
  alert(menuItemText);
}

// Saves a selected style in Document properties, then renumbers headings and updates numbers in links.
// All menu items of submenu 'Select style' (except 'Show current style') run the function
function activateHeadingStyle(obj) {
  // Finds key in headingStyles where value equals obj
  const headingStyle = Object.keys(headingStyles).find(key => headingStyles[key] === obj);

  setDocumentPropertyString("BNumbers_HeadingStyle_Property", headingStyle);
  onOpen();

  doNumberHeadingsAndLinks();
}

// Menu items of submenu 'Select style'
const headingStyles = {
  "numberHeadingsAddChapter23": {
    "name": "H1(~PREFIX~#)/H2-H3/-/figure",
    "separatorAbove": true,
    "xstyle": ['1#', '1', '1', '-', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "overrideH1PrefixWithCustomPrefix": true,
    "numberDepth": 3,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAddChapter234": {
    "name": "H1(~PREFIX~#)/H2-H4/-/figure",
    "xstyle": ['1#', '1', '1', '1', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "overrideH1PrefixWithCustomPrefix": true,
    "numberDepth": 4,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAddChapter2345": {
    "name": "H1(~PREFIX~#)/H2-H5/figure",
    "replacePrefix": true,
    "xstyle": ['1#', '1', '1', '1', '1', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "overrideH1PrefixWithCustomPrefix": true,
    "numberDepth": 5,
    "run": function () { activateHeadingStyle(this); }
  },
    "numberHeadingsAddPrefixAllNumberedAndFigures": {
    "name": "(~PREFIX~)H1-H3/-/-/figure (~PREFIX~)",
    "subMenuBelow": true,
    "xstyle": ['1', '1', '1', '-', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "overrideAllHPrefixWithCustomPrefix": true,
    "numberDepth": 3,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd3Figure": {
    "name": "H1-H3/-/-/figure",
    "separatorAbove": true,
    "xstyle": ['1', '1', '1', '-', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "numberDepth": 3,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd4Figure": {
    "name": "H1-H4/-/figure",
    "xstyle": ['1', '1', '1', '1', '-', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "numberDepth": 4,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAddFigure": {
    "name": "H1-H5/figure",
    "xstyle": ['1', '1', '1', '1', '1', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "numberDepth": 5,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd1WithLinksOrNot": {
    "name": "H1",
    "separatorAbove": true,
    "xstyle": ['1', '-', '-', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 1,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd2WithLinksOrNot": {
    "name": "H1-H2",
    "xstyle": ['1', '1', '-', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 2,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd3WithLinksOrNot": {
    "name": "H1-H3",
    "xstyle": ['1', '1', '1', '-', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 3,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAddPartial4WithLinksOrNot": {
    "name": "H1-H4",
    "xstyle": ['1', '1', '1', '1', '-', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 4,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd5WithLinksOrNot": {
    "name": "H1-H5",
    "xstyle": ['1', '1', '1', '1', '1', '-'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 5,
    "run": function () { activateHeadingStyle(this); }
  },
  "numberHeadingsAdd6WithLinksOrNot": {
    "name": "H1-H6",
    "xstyle": ['1', '1', '1', '1', '1', '1'],
    "prefixlead": [null, null, null, null, null, null],
    "numberDepth": 6,
    "run": function () { activateHeadingStyle(this); }
  }
};

// Menu items of submenu 'Select H1 prefix'
const prefixes = {
  "None": {
    "name": "<none>",
    "func": "setPrefixToNone",
    "value": "",
    "run": function () { activatePrefix(this); }
  },
  "Chapter": {
    "name": "Chapter_",
    "func": "setPrefixToChapter",
    "value": "Chapter ",
    "run": function () { activatePrefix(this); }
  },
  "Section": {
    "name": "Section_",
    "func": "setPrefixToSection",
    "value": "Section ",
    "run": function () { activatePrefix(this); }
  },
  "Session": {
    "name": "Session_",
    "func": "setPrefixToSession",
    "value": "Session ",
    "run": function () { activatePrefix(this); }
  },

  "zeroDot": {
    "name": "0.",
    "func": "setZeroDot",
    "value": "0.",
    "run": function () { activatePrefix(this); }
  },
  "oneDot": {
    "name": "1.",
    "func": "setOneDot",
    "value": "1.",
    "run": function () { activatePrefix(this); }
  },
  "twoDot": {
    "name": "2.",
    "func": "setTwoDot",
    "value": "2.",
    "run": function () { activatePrefix(this); }
  },  
  "threeDot": {
    "name": "3.",
    "func": "setThreeDot",
    "value": "3.",
    "run": function () { activatePrefix(this); }
  },

  "Custom": {
    "name": "<enter>",
    "func": "enterPrefix",
    "run": function () { enterPrefix(); }
  }
};

// All menu items of submenu 'Select H1 prefix' run the function
// Saves a selected prefix in Document properties
function activatePrefix(obj) {
  const prefixStyle = Object.keys(prefixes).find(key => prefixes[key] === obj);
  savePrefixToDocumentProperty(prefixStyle);
}

// Gets custom prefix from user and saves it in Document properties
// Menu item 'Select style' -> 'Select H1 prefix' -> 'Custom'
function enterPrefix() {
  const prefix = getValueFromUser('User defined prefix', 'Please type space after the string, if you want a space before the number.');
  if (prefix != null & prefix != '') {
    setDocumentPropertyString('BNumbers_Custom_Prefix_Property', prefix);
    savePrefixToDocumentProperty('Custom');
  }
}