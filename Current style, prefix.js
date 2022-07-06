function showCurrentStyle() {
  const ui = DocumentApp.getUi();
  const activeStyle = getHeadingStyle();
  onOpen();
  let menuItemText = headingStyles[activeStyle.value]['name'];
  menuItemText = menuItemText.replace('~PREFIX~', activeStyle.prefixText);
  ui.alert(menuItemText);
}

function activateHeadingStyle(obj) {
  // Finds key in headingStyles where value equals obj
  const headingStyle = Object.keys(headingStyles).find(key => headingStyles[key] === obj);

  saveHeadingStyleToDocumentProperty(headingStyle);

  doNumberHeadingsAndLinks();
}

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
    "subMenuBelow": true,
    "xstyle": ['1#', '1', '1', '1', '1', 'figure'],
    "prefixlead": [null, null, null, null, null, 'Figure '],
    "overrideH1PrefixWithCustomPrefix": true,
    "numberDepth": 5,
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
  "Custom": {
    "name": "<enter>",
    "func": "enterPrefix",
    "run": function () { enterPrefix(); }
  }
};

function activatePrefix(obj) {
  const prefixStyle = Object.keys(prefixes).find(key => prefixes[key] === obj);
  savePrefixToDocumentProperty(prefixStyle);
}

function enterPrefix() {
  const prefix = getValueFromUser('User defined prefix', 'Please type space after the string, if you want a space before the number.');
  if (prefix != null & prefix != '') {
    setDocumentPropertyString('BNumbers_Custom_Prefix_Property', prefix);
    savePrefixToDocumentProperty('Custom');
  }
}