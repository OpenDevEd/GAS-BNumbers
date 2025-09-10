// paraSenNumAdd uses the function
function sequentialNumbersX(add, addSen, numberstyle, choice) {
  if (choice == 0) {
    sequentialNumbers(add, addSen, numberstyle);
  } else {
    sequentialNumbersPlus(add, addSen, numberstyle);
  };
};

// Adds text-based markers to paragraphs / sentences
// paraNumAdd, sequentialNumbersX, paraSenNumRemove use the function
function sequentialNumbers(add, addSen, numberstyle) {
  var marker1 = "[";
  var marker2 = "] ";
  var markerS = "-";
  var markerL1 = ".";
  var markerL2 = ".";
  switch (numberstyle) {
    case 0:
      marker1 = '⟦';
      marker2 = '⟧ ';
      break;
    case 1:
      marker1 = '⁅';
      marker2 = '⁆ ';
      break;
    case 2:
      marker1 = '⟦';
      marker2 = '⟧ ';
      markerL1 = "*";
      markerL2 = "*";
      markerS = "+";
      markerL1 = "→";
      markerL2 = "⇉";
      markerS = "⇒";
      break;
    default:
  };
  var myfontsize = null;
  var bgcolor = null;
  var fgcolor = null;
  myfontsize = 6;
  bgcolor = '#ffff00';
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  // use getParagraphs with false: If there's a selection, the selection will be numbered, otherwise the whole doc.
  // var p = getParagraphs(false); -> getParagraphsInBodyAndFootnotesExtended(onePara,getBodyParas,getFootnoteParas) {
  var p = body.getParagraphs();
  var result = "";
  if (numberstyle == -1) {
    result = getValueFromUser('Please enter prefix:');
  };
  var number = -1;
  //DocumentApp.getUi().alert('style: '+numberstyle+ "; ");
  var lead = "";
  var res = "";

  if (add == true) {
    // Paragraphs
    for (var i in p) {
      // Note: With nexted lists the numbering works, but the list style may change from bulleted to numbered.
      var e = p[i];
      var eText = e.getText() + '';
      var eTypeString = e.getHeading() + '';
      if (
        (e.getParent().getType().toString() == 'BODY_SECTION' || e.getParent().getType().toString() == 'TABLE_CELL') &&
        (e.getType().toString() == 'PARAGRAPH' || eTypeString.match(/Heading \d/) || e.getType().toString() == 'LIST_ITEM') &&
        (e.findText("\\S"))
      ) {
        if (add == true) {
          // Add numbers  
          if (eTypeString.match(/Heading (1|2)/)) {
            // DocumentApp.getUi().alert('leadx: '+eTypeString+ "; ");
            var regex = /^[A-Z\~]*?(\d+)\.(\d*)\.? /;
            var found = eText.match(regex);
            if (found) {
              lead = found[1] + markerL1;
              if (found[2]) {
                lead = found[1] + markerL1 + found[2] + markerL2;
              } else {
                lead = found[1] + markerL1;
              };
              //DocumentApp.getUi().alert('lead: '+lead+ "; ");
            } else {
              lead = "";
            };
          } else {
            //    DocumentApp.getUi().alert('leadx: '+e.getType().toString()+ "; ");
          };
          // ADD NUMBERS
          number++;
          //var txt = '[['+number+']] ';
          var txt = "";
          var str = number.toString();
          if (numberstyle == -1) {
            //var str = String.fromCharCode(65+number);
            str = String.fromCharCode(97 + number);
            if (result != "") {
              res = result.toString() + ".";
            };
          };
          txt = marker1 + res + lead + str + marker2;
          var style = e.getAttributes();
          var text = e.getText();
          var placeholderStart = 0;
          var placeholderEnd = 0;
          try {
            var parent = e.getParent();
            //DocumentApp.getUi().alert(e.getType());
            var parPosition = parent.getChildIndex(e);
            var newPara;
            if (e.getType().toString() == 'LIST_ITEM') {
              newPara = parent.insertListItem(parPosition, txt);
            } else {
              newPara = parent.insertParagraph(parPosition, txt);
            };
            newPara.setAttributes(style);
            var thisElementText = newPara.editAsText();
            var highlightuntil = thisElementText.getText().length - 2;
            //if (numberstyle == 0 || numberstyle == 10) {
            if (myfontsize) thisElementText.setFontSize(0, highlightuntil, myfontsize);
            if (fgcolor) thisElementText.setForegroundColor(0, highlightuntil, fgcolor);
            if (bgcolor) thisElementText.setBackgroundColor(0, highlightuntil, bgcolor);
            //} else { // if (style==1) {
            //  thisElementText.setBold(true);
            //};
            var thisElement;
            if (e.getType().toString() == 'LIST_ITEM') {
              // need special treatment for lists:
              var pp2 = parent.getChildIndex(e);
              thisElement = parent.getChild(pp2).merge();
            } else {
              thisElement = e.merge(); // merge these two paragraphs   
            };
            //thisElement.editAsText().insertText(searchResult.getStartOffset(),replacement);
            /* thisElement.asText().insertText(6, "6");
            thisElement.asText().insertText(5, "5");
            thisElement.asText().insertText(4, "4");
            thisElement.asText().insertText(3, "3");
            thisElement.asText().insertText(2, "2");
            thisElement.asText().insertText(1, "1");*/
            // adsen===true
            if (addSen) {
              var sentenceCount = 0;
              // This rexexp only works if the text is contained within one text element. The text may be separated e.g. by a footnotes...
              //var regu = "[\\;\\:\\?\\.] +[A-Z]";
              var regu = "[\\?\\.] +[A-Z]";
              var offset = -1; // offset from the end
              var lengthtally = -1;
              // ... we therefore need to get the children of para, check which are text elements, and operate on each:
              for (var i = 0; i < thisElement.getNumChildren(); i++) {
                var thisSubElement = thisElement.getChild(i);
                if (thisSubElement.getType() == DocumentApp.ElementType.TEXT) {
                  var replacement = "";
                  /* 
                  sentenceCount++;
                  // Place text at the start of the element
                  var cindex = thisElement.getChildIndex(thisSubElement);
                  //replacement = '⟦'+lead+number+'/'+sentenceCount+":"+cindex+'⟧ ' ;  
                  replacement = '⟦'+lead+number+'-'+sentenceCount+'⟧ ' ;  
                  thisSubElement.asText().insertText(0,replacement);
                  var eat = searchResult.getElement().editAsText();
                  if (numberstyle == 0) {
                  if (myfontsize) eat.setFontSize(searchResult.getEndOffsetInclusive()+1+offset, searchResult.getEndOffsetInclusive()+replacement.length+offset,myfontsize);
                  if (fgcolor) eat.setForegroundColor(searchResult.getEndOffsetInclusive()+1+offset, searchResult.getEndOffsetInclusive()+replacement.length+offset,fgcolor);
                  if (bgcolor) eat.setBackgroundColor(searchResult.getEndOffsetInclusive()+1+offset, searchResult.getEndOffsetInclusive()+replacement.length+offset,bgcolor);
                  };
                  */
                  // Now search for other occurences
                  var searchResult = thisSubElement.findText(regu);
                  while (searchResult !== null) {
                    sentenceCount++;
                    // DocumentApp.getUi().alert(thisElement.asText().getText()+', '+searchResult.getStartOffset()+', '+searchResult.getEndOffsetInclusive());
                    // cindex = thisElement.getChildIndex(searchResult.getElement());
                    // replacement = '⟦'+lead+number+'/'+sentenceCount+":"+cindex+'⟧' ; 
                    replacement = marker1 + res + lead + number + markerS + sentenceCount + marker2;
                    searchResult.getElement().asText().insertText(searchResult.getEndOffsetInclusive() + 1 + offset, replacement);
                    //var newElement = searchResult.getElement();
                    //newElement = newElement.replaceText(regu, replacement);      
                    //DocumentApp.getUi().alert('1X: '+thisElement.asText().getText());
                    //try {
                    var eat = searchResult.getElement().editAsText();
                    if (myfontsize) eat.setFontSize(searchResult.getEndOffsetInclusive() + 1 + offset, searchResult.getEndOffsetInclusive() + replacement.length + offset + lengthtally, myfontsize);
                    if (fgcolor) eat.setForegroundColor(searchResult.getEndOffsetInclusive() + 1 + offset, searchResult.getEndOffsetInclusive() + replacement.length + offset + lengthtally, fgcolor);
                    if (bgcolor) eat.setBackgroundColor(searchResult.getEndOffsetInclusive() + 1 + offset, searchResult.getEndOffsetInclusive() + replacement.length + offset + lengthtally, bgcolor);
                    //                  searchResult.getElement().setBackgroundColor(searchResult.getEndOffsetInclusive()+1, searchResult.getEndOffsetInclusive()+1+replacement.length,backgroundColor);
                    //DocumentApp.getUi().alert('2 '+thisElement.asText().getText());
                    // } catch (e) {
                    //  DocumentApp.getUi().alert('exception: '+ e);
                    //};
                    //try {
                    // search for next match
                    searchResult = thisSubElement.findText(regu, searchResult);
                    // } catch (e) {
                    //  DocumentApp.getUi().alert('exception: ' + e);
                    // };
                  };
                } else {
                  // The subelement is not of type TEXT - do nothing. (e.g. footnotes)
                  // DocumentApp.getUi().alert(thisSubElement.getType());
                }
              };
            }
            // End. adsen===true
          } catch (exep) {
            // These should be recorded somehow.
            DocumentApp.getUi().alert('XX-exception: ' + exep);
            // DocumentApp.getUi().alert("Sorry, this feature hasn't been implemented yet.");
            // DocumentApp.getUi().alert(e.getType() + "; " + newPara.getType());
            // DocumentApp.getUi().alert('exception: '+e.getType() + "; " + exep);
          }
          // End. Add numbers
        }
      } else {
        // The element isn't text
        /*
        // REMOVE NUMBERS
        //      var patt1 = new RegExp(/^\[\d+\] /);
        //var patt1 = new RegExp(/./);
        //var rangeElement = e.asText().findText("^\\[\\[\\d+\\]\\] ");
        var rangeElement = e.asText().findText("^⟦\\d+⟧ ");
        if (rangeElement) {
          if (rangeElement.isPartial()) {
            var startOffset = rangeElement.getStartOffset();
            var endOffset = rangeElement.getEndOffsetInclusive();
            rangeElement.getElement().asText().deleteText(startOffset,endOffset);
          } else {
            rangeElement.getElement().removeFromParent();
          }
        }
        // END. REMOVE NUMBERS */
      }
    }
    // End. Paragraphs
  } else {
    // Remove everywhere (para and sentence)
    var removeRegEx = marker1 + "[0-9]+(⇒[0-9]+)?" + marker2 + " ?";
    if (addSen) {
      doc.replaceText(removeRegEx, '');
    } else {
      removeRegEx = "^" + removeRegEx;
      // patt = new RegExp();
      for (var i in p) {
        // Note: With nexted lists the numbering works, but the list style may change from bulleted to numbered.
        var e = p[i];
        e = e.replaceText(removeRegEx, "");
      };
    }
  }
}