function internalHeadingLinksNew(updateNumbers, resetFullLinkText, markInternalHeadingLinks, allHeadingsObj, infoLinksObj) {
  try {
    const doc = DocumentApp.getActiveDocument();
    changeAllBodyLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks);
    changeAllFootnotesLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks);
  }
  catch (error) {
    ui.alert('Error in internalHeadingLinksNew: ' + error);
  }
}

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
    let allHeadingsLastEl = -1;
    let headingText;
    let headingParagraph, headingId;
    requests = [];
    for (let i in bodyElements) {
      headingParagraph = false;
      if (bodyElements[i].paragraph) {
        if (bodyElements[i].paragraph.paragraphStyle) {
          if (bodyElements[i].paragraph.paragraphStyle.headingId) {
            // Logger.log(bodyElements[i].paragraph);
            // Logger.log(bodyElements[i].paragraph.paragraphStyle.headingId);
            headingId = bodyElements[i].paragraph.paragraphStyle.headingId;
            allHeadingParagraphs[headingId] = '';
            allHeadings.push({ headingId: headingId, text: '' });
            allHeadingsLastEl++;
            headingParagraph = true;
          }
        }

        if (bodyElements[i].paragraph.elements) {
          bodyElements[i].paragraph.elements.forEach(function (item) {
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
    }
    //Logger.log(allLinks);
    return { status: 'ok', allHeadingsObj: allHeadingParagraphs, allHeadingsArray: allHeadings };
  }
  catch (error) {
    return { status: 'error', message: 'Error in collectHeadingTexts: ' + error };
  }
}

function changeAllBodyLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks) {
  const element = doc.getBody();
  changeAllLinks(element, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks);
}

function changeAllLinks(element, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks) {

  let text, newText, start, end, indices, partAttributes, numChildren, getIndexFlag, insertMarkerFlag, itsBrokenLink;
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    for (let i = indices.length - 1; i >= 0; i--) {
      partAttributes = element.getAttributes(indices[i]);

      if (partAttributes.LINK_URL) {

        getIndexFlag = false;

        if (/^#heading=/i.test(partAttributes.LINK_URL)) {
          result = /^#heading=(.+)/i.exec(partAttributes.LINK_URL);
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
              if (headingText != text) {
            

                if (resetFullLinkText === true) {
                  newText = headingText;
                  getIndexFlag = true;
                } else {

                  checkLink = /((\d+\.?){1,5})/.exec(text);
                  //Logger.log(checkLink);
                  if (checkLink != null) {
                    linkNumber = checkLink[0];
                    checkHeading = /((\d+\.?){1,5})/.exec(headingText);
                    //Logger.log(checkHeading);
                    if (checkHeading != null) {
                      headingNumer = checkHeading[0];
                      //Logger.log('headingNumer=' + headingNumer + ' linkNumber=' + linkNumber);
                      if (linkNumber != headingNumer) {
                        newText = text.replace(linkNumber, headingNumer);
                        //Logger.log('newText=' + newText + ' text=' + text);
                        getIndexFlag = true;
                      }
                    } else {
                      // infoLinks += allLinks[i].content + ' Heading doesn\'t have number\n';
                    }
                  }


                }
              } else {
                //Logger.log('DONT UPDATE');
              }
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
        changeAllLinks(element.getChild(i), allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks);
      }
    }
  }
}


function changeAllFootnotesLinks(doc, allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks) {
  const footnotes = doc.getFootnotes();
  let footnote, numChildren;
  for (let i in footnotes) {
    footnote = footnotes[i].getFootnoteContents();
    numChildren = footnote.getNumChildren();
    for (let j = 0; j < numChildren; j++) {
      changeAllLinks(footnote.getChild(j), allHeadingsObj, infoLinksObj, updateNumbers, resetFullLinkText, markInternalHeadingLinks);
    }
  }
}