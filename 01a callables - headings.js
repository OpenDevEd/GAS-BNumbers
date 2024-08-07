// Returns debug mode (true or false) 
// Menu_HePaNumbering, doNumberHeadingsAndLinks, numberHeadings use the function
function getDebugMode() {
  let debugMode = false;
  try {
    const debugModeProperty = getDocumentPropertyString("bNumbers_Debug_Mode");
    if (debugModeProperty != null) {
      if (debugModeProperty == "DEBUG_MODE_TRUE") {
        debugMode = true;
      }
    }
  }
  catch (error) {
    //Logger.log(' getDebugMode()) error' + error);
  }
  return debugMode;
}

// Menu 'Turn off debug mode'
function turnOffDebugMode() {
  setDocumentPropertyString("bNumbers_Debug_Mode", "DEBUG_MODE_FALSE");
  onOpen();
}

// 'Turn on debug mode'
function turnOnDebugMode() {
  setDocumentPropertyString("bNumbers_Debug_Mode", "DEBUG_MODE_TRUE");
  onOpen();
}

// Returns object headingStyle that describes current or default style
// doNumberHeadings, removeAllHeadingNumbers, showCurrentStyle, Menu_HePaNumbering use the function
function getHeadingStyle() {
  var headingStyle = {
    value: DEFAULT_STYLE,
    marker: "",
    prefix: DEFAULT_PREFIX,
    prefixMarker: "",
    prefixText: prefixes[DEFAULT_PREFIX]["name"],
    depth: headingStyles[DEFAULT_STYLE]["numberDepth"],
    xstyle: headingStyles[DEFAULT_STYLE]["xstyle"],
    prefixlead: headingStyles[DEFAULT_STYLE]["prefixlead"]
  };
  try {
    var BNumbers_HeadingStyle_Property = getDocumentPropertyString("BNumbers_HeadingStyle_Property");
    if (BNumbers_HeadingStyle_Property != null) {
      headingStyle.value = BNumbers_HeadingStyle_Property;
      headingStyle.marker = "🟢";
      //Logger.log('The Doc has style');
    } else {
      headingStyle.marker = "◯";
    }

    var BNumbers_Prefix_Property = getDocumentPropertyString("BNumbers_Prefix_Property");
    if (BNumbers_Prefix_Property != null) {
      headingStyle.prefix = BNumbers_Prefix_Property;
      headingStyle.prefixMarker = "🟢";
      //Logger.log('The Doc has prefix');

      if (BNumbers_Prefix_Property == 'Custom') {
        //Logger.log('Custom prefix!');
        var BNumbers_Custom_Prefix_Property = getDocumentPropertyString("BNumbers_Custom_Prefix_Property");
        if (BNumbers_Custom_Prefix_Property != null) {
          headingStyle.prefixText = BNumbers_Custom_Prefix_Property;
        }
      } else {
        headingStyle.prefixText = prefixes[BNumbers_Prefix_Property]["value"];
      }

    } else {
      headingStyle.prefixMarker = "◯";
    }
    headingStyle.depth = headingStyles[BNumbers_HeadingStyle_Property]['numberDepth'];
    headingStyle.xstyle = headingStyles[BNumbers_HeadingStyle_Property]['xstyle'];
    headingStyle.prefixlead = headingStyles[BNumbers_HeadingStyle_Property]['prefixlead'];
    if (headingStyles[BNumbers_HeadingStyle_Property]['overrideH1PrefixWithCustomPrefix'] === true) {
      headingStyle.prefixlead[0] = headingStyle.prefixText;
    }
    if (headingStyles[BNumbers_HeadingStyle_Property]['overrideAllHPrefixWithCustomPrefix'] === true) {
      for (style in headingStyle.xstyle) {
        if (headingStyle.xstyle[style] != '-') {
          if (headingStyle.xstyle[style] == 'figure') {
            const regEx = new RegExp(headingStyle.prefixText + '$', 'i');
            if (regEx.test(headingStyle.prefixlead[style]) === false) {
              headingStyle.prefixlead[style] = headingStyle.prefixlead[style] + headingStyle.prefixText;
            }
          } else {
            headingStyle.prefixlead[style] = headingStyle.prefixText;
          }
        }
      }
    }
  }
  catch (error) {
    //Logger.log(' updateStyle() error' + error);
  }
  //Logger.log('get style %s',  headingStyle);
  return headingStyle;
}

// activatePrefix, enterPrefix use the function
function savePrefixToDocumentProperty(prefix) {
  setDocumentPropertyString("BNumbers_Prefix_Property", prefix)
  onOpen();
}

// Menu item Utilities -> 'Add/update heading numbers only [links not updated]'
// Also
// function doNumberHeadingsAndLinks uses the function
// Renumbers headings.
function doNumberHeadings(allHeadingsObj, allHeadingsArray) {

  const headingStyle = getHeadingStyle();

  numberHeadings(true, true, headingStyle.depth, headingStyle.xstyle, headingStyle.prefixlead, null, null, allHeadingsObj, allHeadingsArray);

  if (headingStyle.marker == '◯') {
    setDocumentPropertyString("BNumbers_HeadingStyle_Property", headingStyle.value);
    onOpen();
  }
}

// Menu item 'nha Add/update heading numbers and links'
// activateHeadingStyle, resetFullLinkTextInLinksToHeadings, markInternalHeadingLinks, updateNumbersInLinksToHeadings use the function
// Renumbers headings and updates numbers in texts of links.
function doNumberHeadingsAndLinks(numberHeadings = true, numbersInLinks = true, resetFullLinkText = false, markInternalHeadingLinks = false) {
  /*
  Prior to renumbering, we should check for any broken section links and fix these.
  We can leave that for the future.
  fixBrokenLinksToHeadings();  
  */
  const ui = DocumentApp.getUi();

  const doc = DocumentApp.getActiveDocument();
  const chtResult = collectHeadingTexts(doc);
  if (chtResult.status == 'error') {
    ui.alert(chtResult.message);
    return 0;
  }

  if (numberHeadings) {
    doNumberHeadings(chtResult.allHeadingsObj, chtResult.allHeadingsArray);
  }

  let infoLinks = '';
  const infoLinksObj = { brokenLinkFlag: false, changedLinks: [] };

  internalHeadingLinksNew(numbersInLinks, resetFullLinkText, markInternalHeadingLinks, chtResult.allHeadingsObj, infoLinksObj);

  const debugMode = getDebugMode();

  if (infoLinksObj.changedLinks.length > 0 && debugMode === true) {
    infoLinks = infoLinksObj.changedLinks.join('\n');
  }

  if (infoLinksObj.brokenLinkFlag === true) {
    infoLinks = 'There were broken links. Please search for BROKEN_INTERNAL_LINK_MARKER and fix these. Note that broken links can occur when you press enter at the start of an existing heading.\n\n' + infoLinks;
  }
  if (infoLinks != '') {
    ui.alert(infoLinks);
  }
}

// Menu item Utilities -> 'nh6t Heading 6 tables/figures only'
// Numbers only Headings 6
function updateFigureNumbers(run) {
  // headingStyle = updateFigureNumbers.name;
  // saveHeadingStyleToDocumentProperty(headingStyle);
  //if (run) {
  var xstyle = ['', '', '', '', '', 'figure'];
  const headingStyle = getHeadingStyle();
  numberHeadings(true, false, -1, xstyle, headingStyle.prefixlead);
  //}
};

// Menu item Utilities -> 'nhr Remove all heading numbers'
// Removes all heading numbers of current heading style
function removeAllHeadingNumbers() {
  const headingStyle = getHeadingStyle();
  numberHeadings(false, false, headingStyle.depth, headingStyle.xstyle, headingStyle.prefixlead);
}

