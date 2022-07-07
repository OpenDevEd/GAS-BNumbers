// regexpRestyleOffset, singleReplacePartial use the function
function getParagraphsInBodyAndFootnotesExtended(onePara,getBodyParas,getFootnoteParas) {
  var paraout = [];
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    // If there's a selection, getBodyParas is ignored.
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only modify elements that can be edited as text; skip images and other non-text elements.
      if (element.getElement().editAsText) {
        var elem = element.getElement();
        // var text = elem.editAsText();
        if (elem.getType() == DocumentApp.ElementType.TEXT) {
          elem = elem.getParent();
        }
        if (elem.getType() == DocumentApp.ElementType.PARAGRAPH) {
          var paragraph = elem.asParagraph();
          paraout.push(paragraph);
        } else if (elem.getType() == DocumentApp.ElementType.LIST_ITEM) {
          var paragraph = elem.asListItem();
          paraout.push(paragraph);
        } else {
          DocumentApp.getUi().alert("Cursor is in object that is not paragraph or list item:" + element.getElement().getType() );
        }
      }
    }
  } else {
    if (onePara) {
      // if onePara is true, ignore getBodyParas
      var cursor = DocumentApp.getActiveDocument().getCursor();
      var element = cursor.getElement();
      var paragraph;
      if (element.getParent().getType() == DocumentApp.ElementType.PARAGRAPH) {
        paragraph = element.getParent().asParagraph();
        paraout.push(paragraph);
      } else if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
        paragraph = element.asParagraph();
        paraout.push(paragraph);
      } else if (element.getParent().getType() == DocumentApp.ElementType.LIST_ITEM) {
        paragraph = element.getParent().asListItem();
        paraout.push(paragraph);
      } else if (element.getType() == DocumentApp.ElementType.LIST_ITEM) {
        paragraph = element.asListItem();
        paraout.push(paragraph);
      } else {      
        DocumentApp.getUi().alert("Cursor is in object that is not paragraph or list item: " + element.getParent().getType() + ", parent of " + element.getType());
      }
    } else {
      var doc = DocumentApp.getActiveDocument();
      var paraout = [];
      if (getBodyParas) {
        try {
          var body = doc.getBody();
          paraout = doc.getParagraphs();
        } catch (e) {
          alert("Error in getParagraphsInBodyAndFootnotesExtended: " + e);
        };
      };
      if (getFootnoteParas) {
        var footnote = doc.getFootnotes();
        if (footnote) {
          // alert("Getting fn: "+footnote.length);
          for(var i in footnote){
            if (footnote[i].getFootnoteContents()) {
              var paragraphs = footnote[i].getFootnoteContents().getParagraphs();
              if (paragraphs) {
                //alert("Getting paras: "+paragraphs.length);
                for (var i = 0; i < paragraphs.length; i++) {
                  var element = paragraphs[i];
                  paraout.push(element);
                };
              };
            } else {
              var j = i+1;
              alert("Footnote has no contents. Footnote number= "+j+". This appears to be a GDocs bug that happens if the footnote is suggested text only.");
            };
          }
        } else {
          alert("There are no footnotes.");
        };
      }
    };
  };
  // DocumentApp.getUi().alert(paraout.length);
  // alert(paraout.length);
  return paraout;
}
