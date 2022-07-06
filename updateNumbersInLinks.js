function resetFullLinkTextInLinksToHeadings() {
  doNumberHeadingsAndLinks(numberHeadings = false, numbersInLinks = false, resetFullLinkText = true, markInternalHeadingLinks = false);
}

function markInternalHeadingLinks() {
  doNumberHeadingsAndLinks(numberHeadings = false, numbersInLinks = false, resetFullLinkText = false, markInternalHeadingLinks = true);
}

function updateNumbersInLinksToHeadings() {
  doNumberHeadingsAndLinks(numberHeadings = false, numbersInLinks = true, resetFullLinkText = false, markInternalHeadingLinks = false);
}

function clearInternalLinkMarkers() {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    doc.replaceText(INTERNAL_LINK_MARKER, '');
    doc.replaceText(BROKEN_INTERNAL_LINK_MARKER, '');

    const footnotes = doc.getFootnotes();
    let footnote;
    for (let i in footnotes) {
      footnote = footnotes[i].getFootnoteContents();
      footnote.replaceText(INTERNAL_LINK_MARKER, '');
      footnote.replaceText(BROKEN_INTERNAL_LINK_MARKER, '');
    }
  }
  catch (error) {
    ui.alert('Error in clearInternalLinkMarkers: ' + error);
  }
}

// Out-of-use
const LINK_MARK_STYLE = {
  foregroundColor: {
    color: {
      rgbColor: {
        red: 1.0
      }
    }
  },
  backgroundColor: {
    color: {
      rgbColor: {
        red: 1.0,
        blue: 1.0,
        green: 1.0
      }
    }
  },
  bold: true
};

// Out-of-use
const LINK_MARK_FIELDS = 'foregroundColor, backgroundColor, bold';

// Previous version Update numbers in (hyper)links (to headings)
function internalHeadingLinksOLD(updateNumbers, resetFullLinkText, markInternalHeadingLinks) {
  var ui = DocumentApp.getUi();
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var documentId = doc.getId();
    var document = Docs.Documents.get(documentId);

    var bodyElements = document.body.content;

    var allLinks = [];
    var addLinkFlag = false;
    var allHeadingParagraphs = new Object();
    var headingParagraph, headingId;
    requests = [];
    for (var i in bodyElements) {
      headingParagraph = false;
      if (bodyElements[i].paragraph) {
        if (bodyElements[i].paragraph.paragraphStyle) {
          if (bodyElements[i].paragraph.paragraphStyle.headingId) {
            // Logger.log(bodyElements[i].paragraph);
            // Logger.log(bodyElements[i].paragraph.paragraphStyle.headingId);
            headingId = bodyElements[i].paragraph.paragraphStyle.headingId;
            allHeadingParagraphs[headingId] = '';
            headingParagraph = true;
          }
        }

        if (bodyElements[i].paragraph.elements) {
          bodyElements[i].paragraph.elements.forEach(function (item) {
            if (item.textRun) {
              if (item.textRun.textStyle) {
                if (item.textRun.textStyle.link) {
                  if (item.textRun.textStyle.link.headingId) {
                    allLinks.push({ linkHeadingId: item.textRun.textStyle.link.headingId, content: item.textRun.content, startIndex: item.startIndex, endIndex: item.endIndex });
                  }
                }
              }

              if (item.textRun.content && headingParagraph) {
                // if (item.textRun.content.trim() != '') {
                allHeadingParagraphs[headingId] += item.textRun.content.trim();
                //}
              }
            }
          });
        }
        // When you press enter at the start of an existing heading, headingId still exists, but the paragraph doesn't have content, Doc shows "Heading no longer exists"
        // "Heading no longer exists" case
        if (allHeadingParagraphs[headingId] == '' && headingParagraph) {
          delete allHeadingParagraphs[headingId];
        }
        // End.  "Heading no longer exists" case
      }
    }

    // Logger.log(JSON.stringify(allHeadingParagraphs));
    // Logger.log(allLinks);

    // return 0;

    var checkLink, checkHeading, linkNumber, headingNumer;
    var infoLinks = '', newLinkText, rangeElementOldLink, rangeStart, attr;
    var findOldLink, itsBrokenLink, markInternal, brokenLinkFlag = false, oldLinkFlag = false;
    for (var i = allLinks.length - 1; i >= 0; i--) {
      findOldLink = false;
      itsBrokenLink = false;
      markInternal = false;


      if (markInternalHeadingLinks) {
        findOldLink = true;
        if (!allHeadingParagraphs.hasOwnProperty(allLinks[i].linkHeadingId)) {
          itsBrokenLink = true;
        }
        markInternal = true;
      }

      if (resetFullLinkText) {
        if (allHeadingParagraphs.hasOwnProperty(allLinks[i].linkHeadingId)) {
          if (allHeadingParagraphs[allLinks[i].linkHeadingId] != allLinks[i].content.trim()) {
            newLinkText = allHeadingParagraphs[allLinks[i].linkHeadingId];
            findOldLink = true;
          }
        } else {
          //findOldLink = true;
          itsBrokenLink = true;
        }
      }

      // updateNumbersInLinksToHeadings
      if (updateNumbers) {
        // checkLink = /((\d+\.?){1,5})/.exec(allLinks[i].content);
        // if (checkLink != null) {
        if (allHeadingParagraphs.hasOwnProperty(allLinks[i].linkHeadingId)) {
          checkLink = /((\d+\.?){1,5})/.exec(allLinks[i].content);
          if (checkLink != null) {
            linkNumber = checkLink[0];
            checkHeading = /((\d+\.?){1,5})/.exec(allHeadingParagraphs[allLinks[i].linkHeadingId]);
            if (checkHeading != null) {
              headingNumer = checkHeading[0];
              if (linkNumber != headingNumer) {
                newLinkText = allLinks[i].content.replace(linkNumber, headingNumer);
                findOldLink = true;
              }
            } else {
              infoLinks += allLinks[i].content + ' Heading doesn\'t have number\n';
            }
          }
        } else {
          //findOldLink = true;
          itsBrokenLink = true;
        }

        //        }
      }
      // End. updateNumbersInLinksToHeadings
      if (itsBrokenLink) {
        brokenLinkFlag = true;
      }
      if (findOldLink) {
        oldLinkFlag = true;
      }
      if (itsBrokenLink || findOldLink) {
        infoLinks = addItemToRequestsArray(requests, findOldLink, itsBrokenLink, allLinks[i], newLinkText, infoLinks, markInternal);
      }
    }

    //if (infoLinks == '' && !markInternalHeadingLinks) {
    if (!markInternalHeadingLinks && !oldLinkFlag) {
      if (updateNumbers) {
        infoLinks = 'Links numbering\nunchanged';
      } else if (resetFullLinkText) {
        infoLinks = 'Links texts\nunchanged';
      }
    }

    if (brokenLinkFlag) {
      infoLinks = 'There were broken links. Please search for BROKEN_INTERNAL_LINK_MARKER and fix these. Note that broken links can occur when you press enter at the start of an existing heading.\n\n' + infoLinks;
    }

    if (requests.length > 0) {
      Docs.Documents.batchUpdate({
        requests: requests
      }, documentId);
    }

    if (infoLinks != '') {
      ui.alert(infoLinks);
    }
  }
  catch (error) {
    ui.alert('Error in updateNumbersInLinkedToHeadings. ' + error);
  }
}

// Out-of-use
// internalHeadingLinksOLD uses the function
function addItemToRequestsArray(requests, findOldLink, itsBrokenLink, allLinksItem, newLinkText, infoLinks, markInternal) {
  //if (findOldLink) {
  if (itsBrokenLink) {
    requests.push({
      insertText: {
        text: BROKEN_INTERNAL_LINK_MARKER,
        location: {
          index: allLinksItem.startIndex
        }
      }
    },
      {
        updateTextStyle: {
          range: {
            startIndex: allLinksItem.startIndex,
            endIndex: allLinksItem.startIndex + BROKEN_INTERNAL_LINK_MARKER.length
          },
          text_style: LINK_MARK_STYLE,
          fields: LINK_MARK_FIELDS,
        }
      });
    infoLinks += allLinksItem.content + ' Heading no longer exists\n';
  } else {
    if (markInternal) {
      requests.push({
        insertText: {
          text: INTERNAL_LINK_MARKER,
          location: {
            index: allLinksItem.startIndex
          }
        }
      },
        {
          updateTextStyle: {
            range: {
              startIndex: allLinksItem.startIndex,
              endIndex: allLinksItem.startIndex + INTERNAL_LINK_MARKER.length
            },
            text_style: LINK_MARK_STYLE,
            fields: LINK_MARK_FIELDS,
          }
        });
      //infoLinks += allLinksItem.content + ' marked\n';
    } else {
      requests.push({
        insertText: {
          text: newLinkText,
          location: {
            index: allLinksItem.startIndex + 1
          }
        }
      },
        {
          deleteContentRange: {
            range: {
              startIndex: allLinksItem.startIndex,
              endIndex: allLinksItem.startIndex + 1
            }
          }
        },
        {
          deleteContentRange: {
            range: {
              startIndex: allLinksItem.startIndex + newLinkText.length,
              endIndex: allLinksItem.startIndex + newLinkText.length + allLinksItem.content.length - 1
            }
          }
        });
      infoLinks += allLinksItem.content + ' -> ' + newLinkText + '\n';
    }
  }
  //}
  return infoLinks;
}