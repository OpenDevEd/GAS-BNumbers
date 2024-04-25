///////////// 01 Heading numbers.gs
///////////// Heading numbering
/////////////
// https://stackoverflow.com/questions/39962369/how-to-find-and-remove-blank-paragraphs-in-a-google-document-with-google-apps-sc
// http://stackoverflow.com/questions/12389088/google-docs-drive-number-the-headings
var style = ['1', '1', '1', '1', '1', '1'];

/*
    var xstyle = ['1', '1', '1', '1', '-', 'figure'];

xstyle defines the type of numbering/lettering. E.g. '1' means numerical, and 'A' means letters. 'figure' means: Restart the numbering.
To get numbering like 1.1.A, use
    var xstyle = ['1', '1', 'A', '1', '-', 'figure'];
A "-" means 'no numbering'.

The style 1# = pick-up means that the numbering does not start from 1, but the chapter number is taken from the first heading number:
    var xstyle = ['1#', '1', 'A', '1', '-', 'figure'];
This means that you can start a google doc with the first heading being "Chapter 21." enabling you to have a separate google doc for each chapter.

The prefixlead determines an word appear before the number. For example, this use uses the word 'Figure " for Heading 6:
    var prefixlead = [null, null, null, null, null, 'Figure '];

This uses "Chapter " for heading 1 (and Figure for heading 6):
    var prefixlead = ["Chapter ", null, null, null, null, 'Figure '];

Typically, only heading 1 and heading 6 are customised. However, it would be possible to have this:
  var prefixlead = ["Chapter ", " Section ", " Sub-Section ", null, null, 'Figure '];
to give (a very strange) heading numbering like:
  Chapter 4. Section 3. Sub-Section 1. The title of the heading.

The funnction
    numberHeadings(true, true, 4, xstyle, prefixlead);
Has 5 parameters: whether add or remove
whether to relabel links <- This parameter is now obsolete. Use updateNumbersInLinks instead.
the depth to which to change headings (4)
as well as the styles.

The depth is redundant, because really it's defined through xstyle. However Bjoern was lazy.

*/

function numberHeadingsRepeatH1() {
  var xstyle = ['PUP', '1', '1', '1', '1', 'figure'];
  numberHeadings(true, 6, xstyle);
};


function addMarkupBasedOnStucture(body) {
  var myfontsize = null;
  var bgcolor = null;
  var fgcolor = null;
  myfontsize = 8;
  bgcolor = '#ffff00';
  var p = body.getParagraphs();
  for (var i = 0; i < p.length; i++) {
    var txt = "⁅" + p[i].getParent() + ">" + p[i].getType() + "/" + p[i].getHeading() + "⁆";
    p[i].insertText(0, txt);
    var eat = p[i].editAsText();
    if (myfontsize) eat.setFontSize(0, txt.length - 1, myfontsize);
    if (fgcolor) eat.setForegroundColor(0, txt.length - 1, fgcolor);
    if (bgcolor) eat.setBackgroundColor(0, txt.length - 1, bgcolor);
  };
};

/*
It may be possible to do this differently:

(1) Do not delete heading numbers that are correct. That should speed things up.

(2) Occasionally the merge of paras seems to be a problem... What if new text was inserted after old text, and then old text removed (via .insertText then deleteText).
Where there is no existing number, we could (a) insert from start, and then restyle... or (b) simple insert after the first character and then repeat first char, and then delete it.

*/

function convertListItemsIntoParagraphs(p, headingsButListItems) {
  for (let i in headingsButListItems) {
    var listItem = headingsButListItems[i][0];
    var heading = listItem.getHeading();
    var parent = listItem.getParent();
    var childIndex = listItem.getParent().getChildIndex(listItem);
    const newPar = parent.insertParagraph(childIndex, listItem.getText()).setHeading(heading);
    p[headingsButListItems[i][1]] = newPar;
    listItem.removeFromParent();
  }
}

function numberHeadings(add, changeBodyRefs, maxLevel, numStyle, prefixstr, prefixchar, postfixchar, allHeadingsObj, allHeadingsArray) {

  //Logger.log('add %s, changeBodyRefs  %s, maxLevel %s, numStyle %s, prefixstr %s, prefixchar %s, postfixchar %s', add, changeBodyRefs, maxLevel, numStyle, prefixstr, prefixchar, postfixchar);
  // alert("a="+add+";"+changeBodyRefs+maxLevel+";"+numStyle);
  /*  if (prefixstr) {
      alert("signalling");
    } else {
      alert("none");
    }; */
  // possible improvement: If the heading text is empty, then do not add a heading.
  // Detect the word 'Annex' or Appendix, and change numbering.
  // Existing headsing need to be removed.
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  const p = [];
  var pCheck = doc.getParagraphs();
  var numbers = [0, 0, 0, 0, 0, 0, 0];
  var errors = null;
  var isAppendix = false;

  // The dictionary will track how numbers are renamed.
  var hndict = {};
  var fndict = {};

  var substr = [];

  let ignoredTextRegex;

  for (var i in pCheck) {
    var e = pCheck[i];
    var eText = e.getText() + '';
    var eTypeString = e.getHeading().toString();

    if (!eTypeString.match(/Heading ?\d/i)) {
      // continue if the paragraph is not a heading
      continue;
    }

    eText = eText.trim();

    if (eText == '') {
      // continue if the paragraph is empty
      continue;
    }

    ignoredTextRegex = getIgnoredTextRegex();
    if (ignoredTextRegex !== '') {
      if (ignoredTextRegex.test(eText)) {
        // continue if the heading is matching ignoredTextRegex
        continue;
      }
    }

    // Logger.log('type= ' + pCheck[i].getType());
    p.push(pCheck[i]);
  }

  const headingsButListItems = [];
  const neitherParagraphNorListItem = [];
  const headingsButListItemsTexts = [];
  const neitherParagraphNorListItemTexts = [];
  for (let i = 0; i < p.length; i++) {
    const pType = p[i].getType();
    if (pType === DocumentApp.ElementType.PARAGRAPH) {
      // Perfect
    } else if (pType === DocumentApp.ElementType.LIST_ITEM) {
      headingsButListItems.push([p[i], i]);
      headingsButListItemsTexts.push(p[i].getText());
    } else {
      neitherParagraphNorListItem.push([p[i], i]);
      neitherParagraphNorListItemTexts.push(p[i].getText() + ' ' + pType);
    }
  }

  let wrongHeadings = '';
  let listItemsIssue = false;
  let unknownIssue = false;
  if (headingsButListItems.length > 0) {
    wrongHeadings += '\nList items:\n';
    wrongHeadings += headingsButListItemsTexts.join('\n');
    listItemsIssue = true;
  }
  if (neitherParagraphNorListItem.length > 0) {
    wrongHeadings += '\nOther non-paragraph elements:\n';
    wrongHeadings += neitherParagraphNorListItemTexts.join('\n');
    unknownIssue = true;
  }

  let userResponse;
  if (listItemsIssue || unknownIssue) {
    const addText = unknownIssue ? 'or/and other non-paragraph elements' : '';
    userResponse = getConfirmationFromUser('This document contains headings that are incorrectly formatted as numbered lists ' + addText + '. Would you like to apply proper formatting and numbering?' + wrongHeadings);

    if (userResponse === true) {
      if (listItemsIssue) {
        convertListItemsIntoParagraphs(p, headingsButListItems);
      }
      if (unknownIssue) {
        alert('Fix other non-paragraph elements manually.');
        return 0;
      }
    } else {
      return 0;
    }
  }

  // Go through all paragraphs
  for (var i in p) {
    var e = p[i];
    var eText = e.getText() + '';
    var eTypeString = e.getHeading().toString();
    // alert('HNy '+i+"/"+p.length+" "+eTypeString);

    var patt = new RegExp(/Heading ?(\d)/i);
    var eLevel = patt.exec(eTypeString)[1];   // 1..6 based
    var cLevel = eLevel - 1; //0..5 based
    var currhn = "";

    var figureType = 'Figure';

    if (eLevel <= maxLevel) {
      // ... then remove the heading numbers - see regexp.
      // for removal, the maxLevel is set to 6 - but it can be adjusted to another value, and only those headings will be removed.
      //      var removalPatt = new RegExp(/^[§01-9A-Z\.\-]+ /);
      // Check whether we're in the appendix, and if yes, reset the numbering.
      var re = new RegExp("\\[APPENDIX\\]");
      if (!isAppendix) {
        isAppendix = e.getText().match(re);
        if (isAppendix) {
          numbers[1] = 0;
        };
      }
    };
    if (eLevel <= maxLevel) {
      // alert("H"+eLevel+"<="+maxLevel+"; eLevel-1="+(eLevel-1)+" -> " + numStyle[eLevel-1]+" ");
      // Headings may have prefixes, such as "Chapter". 
      // See if a prefix is part of the style, and if so, match against it.
      // Let's start
      var prefixtext = "";
      if (prefixstr && prefixstr[eLevel - 1]) {
        prefixtext = prefixstr[eLevel - 1];
        // alert(eLevel+", t="+prefixtext);
      } else {
        // alert(eLevel+"oooo, t="+prefixtext);
      };
      var rangeElement = e.asText().findText("^" + prefixtext + "[§01-9A-Z\\.\\-\\_\\~]+[\\.]+[§01-9A-Z\\.\\-]* ");
      // We're now examining the range elements. // THE 'u' flag doesn't work, so cannot match against – or —.
      if (rangeElement) {
        // partial range element:
        if (rangeElement.isPartial()) {
          var startOffset = rangeElement.getStartOffset();
          var endOffset = rangeElement.getEndOffsetInclusive();
          // errors += 'removed: '+ rangeElement.getElement().getText() +", "+ startOffset + "-"+endOffset + "\n";
          // Let's pick up the current number, as we'll need it for various purposes. // This enables picking up a style or number.
          var str = rangeElement.getElement().copy().asText();
          if (endOffset < str.editAsText().getText().length - 1) {
            str = str.deleteText(endOffset, str.editAsText().getText().length - 1);
          };
          if (0 < startOffset - 1) {
            str = str.deleteText(0, startOffset - 1);
          };
          str = str.editAsText().getText();
          str = str.replace(/\. *$/, "");
          var pattx = new RegExp("^" + prefixtext);
          str = str.replace(pattx, "");
          // record the string in the dictionary of heading numbers
          hndict[str] = "";
          currhn = str;
          // alert("currhn="+str);
          // For certain styles, use that number.
          if (numStyle[eLevel - 1] === 'PUP' || numStyle[eLevel - 1] === '1#') {
            if (numStyle[eLevel - 1] === '1#') {
              str = str.replace(/\D/g, "");
              numbers[eLevel] = parseInt(str) - 1;
              numStyle[eLevel - 1] = "1";
            } else {
              substr[eLevel - 1] = str;
            };
            //DocumentApp.getUi().alert('Picked up ' + str);
          };
          rangeElement.getElement().asText().deleteText(startOffset, endOffset);
        } else {
          // complete range element:
          rangeElement.getElement().removeFromParent();
        }
      }
    }

    try {
      if (numStyle[eLevel - 1] === 'figure') {
        //Logger.log('Prefix figure %s ', prefixstr[eLevel - 1]);
        // alert("Hello figure.");
        // ... then remove a figure number / string, irrespective of level, see different regexp below.
        //      errors += "\nfigure: \n";
        // update these both

        var regExpString = "^(" + prefixstr[eLevel - 1] + "|Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) ?(\\-?\\d+(\\.\\d+)?)\\. ?";
        //Logger.log('regExpString %s', regExpString);
        var patt2 = new RegExp(regExpString);
        //        var rangeElement = e.asText().findText("^\⸢(Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) (\\-?\\d+)\\.?\⸥ ?");
        var rangeElement = e.asText().findText(regExpString);
        var figureNumber = "";
        if (rangeElement) {
          if (rangeElement.isPartial()) {
            var startOffset = rangeElement.getStartOffset();
            var endOffset = rangeElement.getEndOffsetInclusive();
            figureType = rangeElement.getElement().getText();
            if (figureType === null) {
              alert("no figure type!");
            } else {
              //          errors +=figureType;
              try {

                var sections = patt2.exec(figureType);
                // alert("Sections: "+sections.length);
                figureType = sections[1];
                figureNumber = sections[2];
              } catch (e) {
                alert("Error figureType: " + e);
              };
              if (figureType === null) {
                alert("no figure type! 2");
              };
            };
            // Delete the numbering:
            rangeElement.getElement().asText().deleteText(startOffset, endOffset);
          } else {
            // This needs thinking through:
            figureNumber = rangeElement.getElement().getText();
            rangeElement.getElement().removeFromParent();
            figureType = rangeElement.getElement().getText();
            //          errors += figureType;
            if (figureType == null) {
              alert("no figure type! 3");
            };
          }
        } else {
          // alert("Hello not figure.");
          // errors+='THis is another use of H6... likely an error in formatting your doc: You requested figure, but H6 is not formatted accordingly';
        };
        if (figureType == null) {
          figureType = "Figure";
        };
        //alert("Fg - "+figureNumber+": "+figureType);
        // record the string in the dictionary of heading numbers
        fndict[figureNumber] = figureNumber;
        currhn = figureNumber;
      };


    } catch (error) {
      alert('Error in figure: ' + error);
    };
    // end of numStyle[eLevel-1] === 'figure'
    // if requested compute new heading numbers
    var txt = '';

    try {
      if (add == true) {
        numbers[eLevel]++;
        if (eLevel == 6 && numStyle[5] === 'figure') {
          // We are at level 6, in a figure.
          // alert("Adding: "+eLevel);
          txt = numbers[1] + "." + numbers[eLevel] + ".";
        } else {
          // Build up the number string for chapters
          for (var l = 1; l <= 6; l++) {
            if (l <= eLevel) {
              // Note that numStyle[l-1] (0-based) corresponds to numbers[l] (1-based)
              // The heading number:
              var ins = numbers[l];
              // The heading number converted to string (according to style):
              var insX = "";
              // Set the numering style based on numStyle, for each heading
              // var numStylel1 = numStyle[l-1];
              txt = txt + getNumberingStyle(l, numbers[l], numStyle, substr);
              if (eLevel == l - 1 && numStyle[l - 1] == 'figure') {
                alert('Figure label now: ' + txt);
                txt = getNumberingStyle(l, numbers[l], numStyle, substr);
              };
            } else {
              // i.e. l > eLevel
              if ((eLevel != 1 && l == 6 && numStyle[5] == 'figure') || numStylel1 === 'T1abcT2' || numStylel1 === 'figure') {
              } else {
                numbers[l] = 0;
              };
            }
            // prefixchar, postfixchar
            var prefixch = "";
            if (prefixchar && prefixchar[eLevel - 1]) {
              prefixch = prefixchar[eLevel - 1];
            }
            txt = prefixch + txt;
            var postfixch = "";
            if (postfixchar && postfixchar[eLevel - 1]) {
              postfixch = postfixchar[eLevel - 1];
            }
            txt = txt + postfixch;
          }
        };
        // done building string
        //Logger.log('eText=' + eText);
        // record value in hcdict, so it can be replaced inreferences
        if (eLevel <= maxLevel || (numStyle[eLevel - 1] === 'figure')) {
          var prefixtext = "";
          if (prefixstr && prefixstr[eLevel - 1]) {
            prefixtext = prefixstr[eLevel - 1];
            // alert(eLevel+", t="+prefixtext);
          } else {
            // alert(eLevel+"oooo, t="+prefixtext);
          };
          // collect the new numbering before perfix is added:
          // if (eLevel <= maxLevel ) @+...
          // if (numStyle[eLevel-1] === 'figure') #+...
          if (numStyle[eLevel - 1] === 'figure') {
            fndict[currhn] = txt;
            fndict[currhn] = fndict[currhn].replace(/\. *$/, "");
            // alert("fndict: currhn="+currhn+"->"+txt);
          } else {
            hndict[currhn] = txt;
            hndict[currhn] = hndict[currhn].replace(/\. *$/, "");
          };
          // alert("currhn="+currhn+"->"+txt);
          txt = prefixtext + txt;
          // record the new number in the hndict
          // insert txt:
          // ... now donw below.
        }
      } else {
        // If add == false, do nothing.
      }
      // We computed the correct number for the heading. SHould it be replaced?
      if (add == true) {
        if (txt != '') {
          if (eLevel <= maxLevel || (numStyle[eLevel - 1] === 'figure')) {
            var style = e.getAttributes();
            var text = e.getText();
            var placeholderStart = 0;
            var placeholderEnd = 0;
            var parent = e.getParent();
            //var d = DocumentApp.getActiveDocument();
            var parPosition = parent.getChildIndex(e);
            // var newPara = d.insertParagraph(parPosition, txt + ' ');
            var newPara = parent.insertParagraph(parPosition, txt + ' ');
            // Logger.log('txt=' + txt);
            // Logger.log('txt eText=' + txt + eText);

            // Update allHeadingsObj allHeadingsArray
            if (allHeadingsObj != null) {
              newHeadingText = txt + ' ' + eText;
              for (let j in allHeadingsArray) {
                //Logger.log(j + 'allHeadingsArray[j].text' + allHeadingsArray[j].text);
                if (allHeadingsArray[j].text == eText) {
                  //Logger.log('Found ' + allHeadingsObj[allHeadingsArray[j].headingId] + ' ->' + txt + eText);
                  allHeadingsObj[allHeadingsArray[j].headingId] = newHeadingText;
                  allHeadingsArray.splice(j, 1);
                  break;
                }
              }
            }
            // End. Update allHeadingsObj

            newPara.setAttributes(style);

            try {
              e.merge(); // merge these two paragraphs   
            } catch (e) {
              // collect error messages, and show at the end
              errors += e + "\n";
            }
          } else {
            //Logger.log("Not replaced, maxLevel=" + maxLevel);
          }
        }
      }
      // now process next para.
    } catch (error) {
      alert("Error in add=true");
    };
  }
  // done going through paragraphs. Were there errors?
  if (errors) {
    DocumentApp.getUi().alert('There were errors.\n' + errors);
  }

  try {
    if (changeBodyRefs) {
      var style = {};
      style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000'; // null
      style[DocumentApp.Attribute.BACKGROUND_COLOR] = null; // null
      style[DocumentApp.Attribute.FONT_SIZE] = 11;
      regexpRestyleOffset("@\\d", style, 0, 1);
      regexpRestyleOffset("#\\d", style, 0, 1);
      // make references safe.
      /* js replace not working
      var regexp = "@([\\d\\.]*\\d)";
      var str = "⁅C$1⁆";
      singleReplace(regexp, str, true, true);      
      regexp = "#([\\d\\.]*\\d)";
      str = "⁅F$1⁆";
      singleReplace(regexp, str, true, true);      
      */
      var map = "";
      var mapf = "";
      // make changes
      var newdict = {};
      for (var key in hndict) {
        var value = hndict[key];
        if (key != value) {
          map += key + " -> " + value + "\n";
          // key = key.replace("\.$","");
          newdict[key] = value;
        };
      };
      if (map == "") {
        map = "unchanged";
      };
      var fnewdict = {};
      for (var key in fndict) {
        var value = fndict[key];
        if (key != value) {
          mapf += key + " -> " + value + "\n";
          // key = key.replace("\.$","");
          fnewdict[key] = value;
        };
      };
      if (mapf == "") {
        mapf = "unchanged";
      };

      const debugMode = getDebugMode();
      if (debugMode === true) {
        DocumentApp.getUi().alert('Chapter numbering\n' + map + '\nFigure numbering\n' + mapf);
      }
      /*
      var arr = Object.keys(hndict);
      arr.sort(function(a, b){
        // ASC  -> a.length - b.length
        // DESC -> b.length - a.length
        return b.length - a.length;
      });
      // DocumentApp.getUi().alert('Chapter numbering\n'+arr.join(", "));
      for(var key in arr) {
        var value = hndict[arr[key]];
        // alert('Chapter numbering\n'+arr[key]+"->"+value);
        if (value && arr[key]) {
          if (arr[key] != '' && arr[key] != value) {
            // alert('Chapter numbering x\n'+arr[key]+"->"+value);
            //singleReplace("@"+arr[key],"[["+arr[key]+"]][:]"+value,false,false,null);
            try {
              singleReplace("@"+arr[key],"[:]"+value,false,false,null);
            } catch (e) {
              alert("Error @->[:] - "+e);
            };
          };
        };
        */
      var arr = Object.keys(newdict);
      for (var key in arr) {
        if (arr[key] != '') {
          var regexp = "@" + arr[key];
          var str = "⁅C" + arr[key] + "⁆";
          //alert(regexp+ " -> " + str);
          singleReplace(regexp, str, false, false);
        };
      };
      for (var key in arr) {
        var value = newdict[arr[key]];
        if (value && arr[key]) {
          if (arr[key] != '' && arr[key] != value) {
            // alert('Chapter numbering x\n'+arr[key]+"->"+value);
            //singleReplace("@"+arr[key],"[["+arr[key]+"]][:]"+value,false,false,null);
            try {
              singleReplace("⁅C" + arr[key] + "⁆", "@" + value, false, false, null);
            } catch (e) {
              alert("Error @->[:] - " + e);
            };
          };
        };
      };
      var farr = Object.keys(fnewdict);
      for (var key in farr) {
        if (farr[key] != '') {
          var regexp = "#" + farr[key];
          var str = "⁅F" + farr[key] + "⁆";
          //alert("Replace: "+regexp+ " -> "+str);
          singleReplace(regexp, str, false, false);
        };
      };
      for (var key in farr) {
        var value = fnewdict[farr[key]];
        if (value && farr[key]) {
          if (farr[key] != '' && farr[key] != value) {
            // alert('Chapter numbering x\n'+arr[key]+"->"+value);
            //singleReplace("@"+arr[key],"[["+arr[key]+"]][:]"+value,false,false,null);
            try {
              //alert("Replace: F"+farr[key]+ " -> #"+value);
              singleReplace("⁅F" + farr[key] + "⁆", "#" + value, false, false, null);
            } catch (e) {
              alert("Error #->[:] - " + e);
            };
          };
        };

      };
      style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666'; // null
      style[DocumentApp.Attribute.BACKGROUND_COLOR] = null; // null
      style[DocumentApp.Attribute.FONT_SIZE] = 6;
      //regexpRestyleOffset("@\\d",style,0,1);
      //regexpRestyleOffset("#\\d",style,0,1);  
      try {
        singleReplace("[:]", "@", false, false, null);
        // reinstate references
        /* singleReplace("⁅C","@", false, false);      
        singleReplace("⁅F","#", false, false);      
        singleReplace("⁆","", false, false);            */
      } catch (e) {
        alert("Error [:]->@ - " + e);
      };
    }
  } catch (e) {
    alert("Error in changeBodyRefs: " + e);
  };
}

// Function numberHeadings uses the function
function getNumberingStyle(l, ins, numStyle, substr) {
  var figureType = "Figure";
  numStylel1 = numStyle[l - 1];
  var insX = "";
  var txt = "";
  if (numStylel1 === 'A') {
    insX = String.fromCharCode(64 + ins);
  } else if (numStylel1 === '1') {
    insX = ins.toString();
  } else if (numStylel1 === 'S1') {
    insX = "S" + ins.toString();
  } else if (numStylel1 === '§1') {
    insX = "§" + ins.toString();
  } else if (numStylel1 === 'A.1') {
    insX = "A." + ins.toString();
  } else if (numStylel1 === 'B.1') {
    insX = "B." + ins.toString();
  } else if (numStylel1 === 'C.1') {
    insX = "C." + ins.toString();
  } else if (numStylel1 === 'A1') {
    insX = "A" + ins.toString();
  } else if (numStylel1 === 'B1') {
    insX = "B" + ins.toString();
  } else if (numStylel1 === 'C1') {
    insX = "C" + ins.toString();
  } else if (numStylel1 === 'A-1') {
    insX = "A-" + ins.toString();
  } else if (numStylel1 === 'B-1') {
    insX = "B-" + ins.toString();
  } else if (numStylel1 === 'C-1') {
    insX = "C-" + ins.toString();
  } else if (numStylel1 === 'A1-1') {
    insX = "A1-" + ins.toString();
  } else if (numStylel1 === 'B1-1') {
    insX = "B1-" + ins.toString();
  } else if (numStylel1 === 'B2-1') {
    insX = "B2-" + ins.toString();
  } else if (numStylel1 === 'C1-1') {
    insX = "C" + ins.toString();
  } else if (numStylel1 === 'B1-') {
    insX = "B" + ins.toString();
  } else if (numStylel1 === 'B1-') {
    insX = "B" + ins.toString();
  } else if (numStylel1 === '§1-') {
    insX = "§" + ins.toString();
  } else if (numStylel1 === '[1]') {
    insX = "[" + ins.toString() + "]";
  } else if (numStylel1 === 'Teil1') {
    insX = ins.toString();
    if (eLevel == 1) {
      if (isAppendix) {
        insX = "APPENDIX~" + ins.toString();
      } else {
        insX = "KAPITEL~" + ins.toString();
      };
    };
  } else if (numStylel1 === 'T1abcT2') {
    // This style has an odd numbering - used in DFID proposal.
    ins--;
    switch (ins) {
      case -1:
        insX = "T00";
        break;
      case 0:
        insX = "T0";
        break;
      case 1:
        insX = "T1A";
        break;
      case 2:
        insX = "T1B";
        break;
      case 3:
        insX = "T1C";
        break;
      default:
        var insm = ins - 2;
        insX = "T" + insm.toString()
    }
  } else if (numStylel1 === 'figure') {
    if (figureType == null) {
      alert("no figure type! 10");
      figureType = "Figure";
    };
    // insX = "⸢"+figureType+ " " + ins.toString() + ".⸥"; 
    insX = "" + figureType + " " + ins.toString() + ".";
  } else if (numStylel1 === 'PUP') {
    insX = substr[l - 1];
  } else {
    insX = "X" + ins.toString();
  };
  // Special treatment of 'A1-' styles
  if (l - 2 >= 0) {
    switch (numStyle[l - 2]) {
      case "A1-":
      case "B1-":
      case "C1-":
        var regex = /\.$/gi;
        txt = txt.replace(regex, '');
        //                txt += '–';
        txt += '-';
        break;
      default:
        break;
    };
  };
  // Add to txt (I.e. assemble the full heading)
  switch (numStylel1) {
    case '§1-':
      txt += insX + '-';
      break;
    /*    case "A1-":
    case "B1-":
    case "C1-":
    txt += insX + '.' ;
    break; */
    case "figure":
      // Do not append, but replace XXX
      if (numStylel1 == 'figure') {
        txt = insX;
      } else {
        txt = '';
      };
      break;
    default:
      txt += insX + '.';
  }
  return txt;
}

// Menu item Utilities -> 'nhprefix prefix all headings with a string'
// Gets a string from user, adds the string at the beginning of every headings.
function prefixHeadings() {
  var prefix = getValueFromUser("All headings will be prefixed with a fixed string. Please enter the string.");
  if (prefix != "") {
    var doc = DocumentApp.getActiveDocument();
    var p = doc.getParagraphs();
    for (var i in p) {
      var e = p[i];
      var eText = e.getText() + '';
      var eTypeString = e.getHeading() + '';

      /*
      New Docs return HEADING1
      Old Docs return Heading 1
      RegEx
      /Heading \d/
      was changed to
      /Heading ?\d/i
      */
      if (!eTypeString.match(/Heading ?\d/i)) {
        // continue if the paragraph is not a heading
        continue;
      }
      e.editAsText().insertText(0, prefix);
    }
  };
}