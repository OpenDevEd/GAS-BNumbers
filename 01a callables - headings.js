var style = ['1', '1', '1', '1', '1', '1'];

// Default:
var headingStyle = getDocumentPropertyString("BNumbers_HeadingStyle_Property") || "numberHeadingsAdd3Figure";

console.log("Default style: " + headingStyle);


function saveHeadingStyleToDocumentProperty(headingstyle) {
  setDocumentPropertyString("BNumbers_HeadingStyle_Property", headingstyle)
  console.log("Setting style")
}

function getHeadingStyleToDocumentProperty() {
  return getDocumentPropertyString("BNumbers_HeadingStyle_Property")
}

function doNumberHeadings() {
  eval(headingStyle)(true);
}

function doNumberHeadingsAndLinks() {
  /*
  Prior to renumbering, we should check for any broken section links and fix these.
  We can leave that for the future.
  fixBrokenLinksToHeadings();  
  */
  doNumberHeadings();

  DocumentApp.getActiveDocument().saveAndClose();
  
  updateNumbersInLinksToHeadings();
}

function numberHeadingsAdd6(run) {
  headingStyle = numberHeadingsAdd6.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    numberHeadings(true, false, 6, style);
  }
}

function numberHeadingsAddPartial3(run) {
  headingStyle = numberHeadingsAddPartial3.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    numberHeadings(true, false, 3, style);
  }
}

function numberHeadingsAddPartial4(run) {
  headingStyle = numberHeadingsAddPartial4.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    numberHeadings(true, false, 4, style);
  }
}

function numberHeadingsAddFigure(run) {
  headingStyle = numberHeadingsAddFigure.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1', '1', '1', '1', '1', 'figure'];
    var prefixlead = [null, null, null, null, null, 'Figure '];
    numberHeadings(true, true, 4, xstyle, prefixlead);
  }
}

function numberHeadingsAdd4Figure(run) {
  headingStyle = numberHeadingsAdd4Figure.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1', '1', '1', '1', '-', 'figure'];
    var prefixlead = [null, null, null, null, null, 'Figure '];
    numberHeadings(true, true, 4, xstyle, prefixlead);
  }
}

// 'H1-3 (update links)'
function numberHeadingsAdd3WithLinks(run) {
  headingStyle = numberHeadingsAdd3Figure.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1', '1', '1', '-', '-', '-'];
    var prefixlead = [null, null, null, null, null, null];
    numberHeadings(true, true, 3, xstyle, prefixlead);
  }
}

// 'H1-3/-/-/figure (update links)'
function numberHeadingsAdd3Figure(run) {
  headingStyle = numberHeadingsAdd3Figure.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1', '1', '1', '-', '-', 'figure'];
    var prefixlead = [null, null, null, null, null, 'Figure '];
    numberHeadings(true, true, 3, xstyle, prefixlead);
  }
}

function updateFigureNumbers(run) {
  headingStyle = updateFigureNumbers.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['', '', '', '', '', 'figure'];
    numberHeadings(true, false, -1, xstyle);
  }
};

// OBSOLETE
// Govet current with 'pickup': First heading number is determined from first H1 heading.
function numberHeadingsAddTeilFigure11pickGerman(run) {
  headingStyle = numberHeadingsAddTeilFigure11pickGerman.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '1', '-', '-'];
    var prefixlead = ['Kapitel ', null, null, null, null, 'Abbildung '];
    //  var prefixchar = ['[','{','(','<','',''];
    //  var postfixcar = [']','}',')','>','',''];
    //  numberHeadings(true,false,4,xstyle, prefixlead, prefixchar,postfixcar );
    numberHeadings(true, true, 4, xstyle, prefixlead);
  }
}

// First heading number is determined from first H1 heading.
//function numberHeadingsAddTeilFigure11pick(run) {
function numberHeadingsAddChapter234() {
  headingStyle = numberHeadingsAddChapter234.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '1', '-', '-'];
    // Elena:
    var prefixlead = [SELECTED_PREFIX, null, null, null, null, 'Figure '];
    numberHeadings(true, true, 4, xstyle, prefixlead);
  }
}

function numberHeadingsAddChapter23(run) {
  headingStyle = numberHeadingsAddChapter23.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '-', '-', '-'];
    var prefixlead = [SELECTED_PREFIX, null, null, null, null, 'Figure '];
    numberHeadings(true, true, 3, xstyle, prefixlead);
  }
}

// OBSOLETE
function numberHeadingsAddSection23(run) {
  headingStyle = numberHeadingsAddSection23.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '-', '-', '-'];
    var prefixlead = ['Section ', null, null, null, null, 'Figure '];
    numberHeadings(true, true, 3, xstyle, prefixlead);
  }
}

// OBSOLETE
function numberHeadingsAddSession23(run) {
  headingStyle = numberHeadingsAddSession23.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '-', '-', '-'];
    var prefixlead = ['Session ', null, null, null, null, 'Figure '];
    numberHeadings(true, true, 3, xstyle, prefixlead);
  }
}

// OBSOLETE
function numberHeadingsAddSection234(run) {
  headingStyle = numberHeadingsAddSection234.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['1#', '1', '1', '1', '-', '-'];
    var prefixlead = ['Section ', null, null, null, null, 'Figure '];
    numberHeadings(true, true, 4, xstyle, prefixlead);
  }
}



/*
// ?
function numberHeadingsAddPartial1A11(run) {
  headingStyle = numberHeadingsAddPartial1A11.name;
  saveHeadingStyleToDocumentProperty(headingStyle);
  if (run) {
    var xstyle = ['§1-', 'A', '1', '1', '1', '1'];
    numberHeadings(true, false, 4, xstyle);
  }
}

// This one is used for DFID.
function numberHeadingsAddPartialB11T(){
  var xstyle = ['B1-1','T1abcT2','1','1','1','figure'];
  numberHeadings(true,false,4,xstyle);
}

function numberHeadingsAddPartialB1T(){
  var xstyle = ['B1-','T1abcT2','1','1','1','figure'];
  numberHeadings(true,false,4,xstyle);
}

// DFID
function numberHeadingsAddPartialBdash1T(){
  var xstyle = ['B-1','T1abcT2','1','1','1','1'];
  numberHeadings(true,false,4,xstyle);
}

// DFID
function numberHeadingsAddPartialBdot1T(){
  var xstyle = ['B.1','T1abcT2','1','1','1','1'];
  numberHeadings(true,false,4,xstyle);
}
*/

// General
function numberHeadingsClear6() {
  numberHeadings(false, false, 6, style);
}

function numberHeadingsClear3() {
  numberHeadings(false, false, 3, style);
}

function numberHeadingsClear4() {
  numberHeadings(false, false, 4, style);
}

