// Returns object headingStyle that describes current or default style
// doNumberHeadings, removeAllHeadingNumbers, showCurrentStyle, Menu_HePaNumbering use the function
function getHeadingStyle() {
  var headingStyle = {
    value: DEFAULT_STYLE,
    marker: "",
    prefix: DEFAULT_PREFIX,
    prefixMarker: "",
    prefixText: prefixes[DEFAULT_PREFIX]["name"],
    depth: headingStyles[DEFAULT_STYLE]["depth"],
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
  }
  catch (error) {
    //Logger.log(' updateStyle() error' + error);
  }
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

  if (infoLinksObj.changedLinks.length > 0) {
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
  numberHeadings(true, false, -1, xstyle);
  //}
};

// Menu item Utilities -> 'nhr Remove all heading numbers'
// Removes all heading numbers of current heading style
function removeAllHeadingNumbers() {
  const headingStyle = getHeadingStyle();
  numberHeadings(false, false, headingStyle.depth, headingStyle.xstyle);
}

