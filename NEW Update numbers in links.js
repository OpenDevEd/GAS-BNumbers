// doNumberHeadingsAndLinks use the function
function internalHeadingLinksNew(updateNumbers, resetFullLinkText, markInternalHeadingLinks, allHeadingsObj, infoLinksObj) {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const { style: withWithoutDots } = getSettings(true, 'withoutDots', withWithoutDotsStyles, 'WITH_WITHOUT_DOTS_SETTINGS');
    const withoutDots = withWithoutDots === 'withoutDots' ? true : false;
    changeAllBodyLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots);
    changeAllFootnotesLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots);
  }
  catch (error) {
    ui.alert('Error in internalHeadingLinksNew: ' + error);
  }
}

// doNumberHeadingsAndLinks use the function
// Collects texts of all headings
// Returns object allHeadingParagraphs 
// {headingId: 'Heading text', ...}
// and array allHeadings
// [{ headingId: headingId, text: 'Heading text' }, ...]
function collectHeadingTexts(doc) {
  try {
    if (!doc) {
      doc = DocumentApp.getActiveDocument();
    }

    const documentId = doc.getId();
    const document = Docs.Documents.get(documentId);

    const bodyElements = document.body.content;

    const allLinks = [];
    const allHeadingParagraphs = new Object();
    const allHeadings = [];
    requests = [];
    for (let i in bodyElements) {
      // If body element contains table
      if (bodyElements[i].table) {
        if (bodyElements[i].table.tableRows) {
          for (let j in bodyElements[i].table.tableRows) {
            if (bodyElements[i].table.tableRows[j].tableCells) {
              for (let k in bodyElements[i].table.tableRows[j].tableCells) {
                if (bodyElements[i].table.tableRows[j].tableCells[k].content) {
                  for (let l in bodyElements[i].table.tableRows[j].tableCells[k].content) {
                    if (bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph) {
                      collectHeadingTextsHelper(bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph, allHeadingParagraphs, allHeadings, allLinks);
                    }
                  }
                }
              }
            }
          }

        }
      }
      // End. If body element contains table

      // If body element contains paragraph
      if (bodyElements[i].paragraph) {
        collectHeadingTextsHelper(bodyElements[i].paragraph, allHeadingParagraphs, allHeadings, allLinks);
      }
      // End. If body element contains paragraph
    }
    //Logger.log(allLinks);
    return { status: 'ok', allHeadingsObj: allHeadingParagraphs, allHeadingsArray: allHeadings };
  }
  catch (error) {
    if (error.toString().includes('API call to docs.documents.get failed with error: Internal error encountered.')) {
      return { status: 'error', message: 'Sorry, bNumbers has encountered an error. It’s a known error we haven’t been able to trace yet. However, you can fix it by copying the contents of your document to a new document and running bNumbers in the new document. (Error in collectHeadingTexts: Error in collectHeadingTexts: ' + error + ')' };
    } else {
      return { status: 'error', message: 'Error in collectHeadingTexts: ' + error };
    }
  }
}

// Checks paragraphs in body and paragraphs in tables
function collectHeadingTextsHelper(paragraph, allHeadingParagraphs, allHeadings, allLinks) {
  let headingParagraph = false;
  let headingText, headingId;
  if (paragraph.paragraphStyle) {
    if (paragraph.paragraphStyle.headingId) {
      // Logger.log(paragraph);
      // Logger.log(paragraph.paragraphStyle.headingId);
      headingId = paragraph.paragraphStyle.headingId;
      allHeadingParagraphs[headingId] = '';
      allHeadings.push({ headingId: headingId, text: '' });
      allHeadingsLastEl = allHeadings.length - 1;
      headingParagraph = true;
    }
  }

  if (paragraph.elements) {
    paragraph.elements.forEach(function (item) {
      if (item.textRun) {
        // Collects links
        if (item.textRun.textStyle) {
          if (item.textRun.textStyle.link) {
            if (item.textRun.textStyle.link.headingId) {
              allLinks.push({ linkHeadingId: item.textRun.textStyle.link.headingId, content: item.textRun.content, startIndex: item.startIndex, endIndex: item.endIndex });
            }
          }
        }
        // End. Collects links

        if (item.textRun.content && headingParagraph) {
          // if (item.textRun.content.trim() != '') {
          headingText = item.textRun.content.replace('\n', '');
          allHeadingParagraphs[headingId] += headingText;
          allHeadings[allHeadingsLastEl]['text'] += headingText;
          //}
        }
      }
    });
  }
  // When you press enter at the start of an existing heading, headingId still exists, but the paragraph doesn't have content, Doc shows "Heading no longer exists"
  // "Heading no longer exists" case
  if (allHeadingParagraphs[headingId] == '' && headingParagraph) {
    delete allHeadingParagraphs[headingId];
    allHeadings.pop();
    allHeadingsLastEl--;
  }
  // End.  "Heading no longer exists" case


}


// internalHeadingLinksNew uses the function
function changeAllBodyLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots) {
  const element = doc.getBody();
  changeAllLinks(element, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots);
}

// internalHeadingLinksNew uses the function
function changeAllFootnotesLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots) {
  const footnotes = doc.getFootnotes();
  let footnote, numChildren;
  for (let i in footnotes) {
    footnote = footnotes[i].getFootnoteContents();
    numChildren = footnote.getNumChildren();
    for (let j = 0; j < numChildren; j++) {
      changeAllLinks(footnote.getChild(j), allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots);
    }
  }
}

// changeAllBodyLinks, changeAllFootnotesLinks use the function
function changeAllLinks(element, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots) {

  let text, newText, start, end, indices, partAttributes, numChildren, getIndexFlag, insertMarkerFlag, itsBrokenLink;
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    for (let i = indices.length - 1; i >= 0; i--) {
      partAttributes = element.getAttributes(indices[i]);

      if (partAttributes.LINK_URL) {
        getIndexFlag = false;

        if (/^[?#].*heading=/i.test(partAttributes.LINK_URL)) {
          result = /#heading=(.+)/i.exec(partAttributes.LINK_URL);
          headingIdLink = result[1];
          //Logger.log('internal headingIdLink=' + headingIdLink);

          start = indices[i];
          elementText = element.getText();
          if (i == indices.length - 1) {
            end = element.getText().length - 1;
          } else {
            end = indices[i + 1] - 1;
          }
          text = elementText.substr(start, end - start + 1);
          //Logger.log('text=' + text);

          if (allHeadingsObj.hasOwnProperty(headingIdLink) === false) {
            itsBrokenLink = true;
            getIndexFlag = true;
          } else {
            if (updateNumbers === true || resetFullLinkText === true) {
              headingText = allHeadingsObj[headingIdLink];
              // if (headingText != text) {


                if (resetFullLinkText === true) {
                  newText = headingText;
                  getIndexFlag = true;
                } else {

                  checkLink = /((\d+\.?){1,5})/.exec(text);
                  // Logger.log(checkLink);
                  if (checkLink != null) {
                    linkNumber = checkLink[0];
                    checkHeading = /((\d+\.?){1,5})/.exec(headingText);
                    // Logger.log(checkHeading);
                    if (checkHeading != null) {
                      headingNumer = checkHeading[0];
                      // Logger.log('headingNumer=' + headingNumer + ' linkNumber=' + linkNumber);

                      let targetHeadingNumber;
                      if (headingNumer.charAt(headingNumer.length - 1) === '.') {
                        // targetHeadingNumber = headingNumer.slice(0, -1);
                        targetHeadingNumber = withoutDots === true ? headingNumer.slice(0, -1) : headingNumer;
                      } else {
                        // targetHeadingNumber = headingNumer;
                        targetHeadingNumber = withoutDots === true ? headingNumer : headingNumer + '.';
                      }
                      if (linkNumber != targetHeadingNumber) {
                        newText = text.replace(linkNumber, targetHeadingNumber);
                        // Logger.log('newText=' + newText + ' text=' + text);
                        getIndexFlag = true;
                      }
                    } else {
                      // infoLinks += allLinks[i].content + ' Heading doesn\'t have number\n';
                    }
                  }


                }
              // } else {
              //   //Logger.log('DONT UPDATE');
              // }
            } else if (markInternalHeadingLinks === true) {
              getIndexFlag = true;
            }
          }

        } else {
          //Logger.log('External link');
        }

        if (getIndexFlag === true) {
          insertMarkerFlag = false;
          if (itsBrokenLink === true || markInternalHeadingLinks === true) {
            insertMarkerFlag = true;
            if (itsBrokenLink === true) {
              infoLinksObj.brokenLinkFlag = true;
              marker = BROKEN_INTERNAL_LINK_MARKER;
            } else {
              marker = INTERNAL_LINK_MARKER;
            }
          } else {
            element.deleteText(start, end);
            element.insertText(start, newText).setAttributes(start, start + newText.length - 1, partAttributes);
            infoLinksObj.changedLinks.push(text + ' -> ' + newText);
          }

          if (insertMarkerFlag === true) {
            element.insertText(indices[i], marker).setLinkUrl(indices[i], indices[i] + marker.length - 1, null).setAttributes(indices[i], indices[i] + marker.length - 1, LINK_MARK_STYLE_NEW);
          }
        }
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        changeAllLinks(element.getChild(i), allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks, withoutDots);
      }
    }
  }
}