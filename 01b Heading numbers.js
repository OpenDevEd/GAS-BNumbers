///////////// 01 Heading numbers.gs
///////////// Heading numbering
/////////////
// https://stackoverflow.com/questions/39962369/how-to-find-and-remove-blank-paragraphs-in-a-google-document-with-google-apps-sc
// http://stackoverflow.com/questions/12389088/google-docs-drive-number-the-headings
var style = ['1','1','1','1','1','1'];

function numberHeadingsRepeatH1() {
  var xstyle = ['PUP','1','1','1','1','figure'];
  numberHeadings(true,6,xstyle);
};


function addMarkupBasedOnStucture(body) {
  var myfontsize = null;
  var bgcolor = null;
  var fgcolor = null;
  myfontsize = 8;
  bgcolor = '#ffff00';
  var p = body.getParagraphs();
  for (var i=0; i<p.length; i++) {
    var txt = "⁅"+p[i].getParent() + ">" + p[i].getType() + "/" + p[i].getHeading()+ "⁆";
    p[i].insertText(0,txt);
    var eat = p[i].editAsText();
    if (myfontsize) eat.setFontSize(0,txt.length-1,myfontsize);
    if (fgcolor) eat.setForegroundColor(0,txt.length-1,fgcolor);
    if (bgcolor) eat.setBackgroundColor(0,txt.length-1,bgcolor);
  };
};

/*
It may be possible to do this differently:

(1) Do not delete heading numbers that are correct. That should speed things up.

(2) Occasionally the merge of paras seems to be a problem... What if new text was inserted after old text, and then old text removed (via .insertText then deleteText).
Where there is no existing number, we could (a) insert from start, and then restyle... or (b) simple insert after the first character and then repeat first char, and then delete it.

*/
function numberHeadings(add, changeBodyRefs, maxLevel, numStyle, prefixstr, prefixchar, postfixchar){
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
  var p = doc.getParagraphs();
  var numbers = [0,0,0,0,0,0,0];
  var errors = null;
  var isAppendix = false;
  
  // The dictionary will track how numbers are renamed.
  var hndict = {};
  var fndict = {};
  
  var substr = [];
  
  
  // Go through all paragraphs
  for (var i in p) {
    var e = p[i];
    var eText = e.getText()+'';
    var eTypeString = e.getHeading().toString();
    // alert('HNy '+i+"/"+p.length+" "+eTypeString);
    if (!eTypeString.match(/Heading ?\d/i)) {
      // continue if the paragraph is not a heading
      continue;
    }

    var patt = new RegExp(/Heading ?(\d)/i);
    var eLevel = patt.exec(eTypeString)[1];   // 1..6 based
    var cLevel = eLevel-1; //0..5 based
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
      if (prefixstr && prefixstr[eLevel-1]) {
        prefixtext = prefixstr[eLevel-1] ;
        // alert(eLevel+", t="+prefixtext);
      } else {
        // alert(eLevel+"oooo, t="+prefixtext);
      };
      var rangeElement = e.asText().findText("^"+prefixtext+"[§01-9A-Z\\.\\-\\_\\~]+[\\.]+[§01-9A-Z\\.\\-]* ");      
      // We're now examining the range elements. // THE 'u' flag doesn't work, so cannot match against – or —.
      if (rangeElement) {
        // partial range element:
        if (rangeElement.isPartial()) {
          var startOffset = rangeElement.getStartOffset();
          var endOffset = rangeElement.getEndOffsetInclusive();
          // errors += 'removed: '+ rangeElement.getElement().getText() +", "+ startOffset + "-"+endOffset + "\n";
          // Let's pick up the current number, as we'll need it for various purposes. // This enables picking up a style or number.
          var str = rangeElement.getElement().copy().asText();
          if (endOffset < str.editAsText().getText().length-1) {
            str = str.deleteText(endOffset,str.editAsText().getText().length-1);
          };
          if (0 < startOffset-1) {
            str = str.deleteText(0,startOffset-1);
          };
          str = str.editAsText().getText();
          str = str.replace(/\. *$/,"");
          var pattx = new RegExp("^"+prefixtext);
          str = str.replace(pattx,"");
          // record the string in the dictionary of heading numbers
          hndict[str] = "";
          currhn = str;
          // alert("currhn="+str);
          // For certain styles, use that number.
          if (numStyle[eLevel-1] === 'PUP' || numStyle[eLevel-1] === '1#') {
            if (numStyle[eLevel-1] === '1#') {
              str = str.replace(/\D/g,"");
              numbers[eLevel] = parseInt(str)-1;
              numStyle[eLevel-1] = "1";
            } else {
              substr[eLevel-1] = str;            
            };
            //DocumentApp.getUi().alert('Picked up ' + str);
          };
          rangeElement.getElement().asText().deleteText(startOffset,endOffset);         
        } else {
          // complete range element:
          rangeElement.getElement().removeFromParent();
        }
      }
    }
    try {            
      if (numStyle[eLevel-1] === 'figure') {
        // alert("Hello figure.");
        // ... then remove a figure number / string, irrespective of level, see different regexp below.
        //      errors += "\nfigure: \n";
        // update these both
//        var patt2 = new RegExp(/^\⸢(Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) (\-?\d+)\.\⸥ ?/);
        var patt2 = new RegExp(/^(Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) (\-?\d+(\.\d+)?)\. ?/);
//        var rangeElement = e.asText().findText("^\⸢(Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) (\\-?\\d+)\\.?\⸥ ?");      
        var rangeElement = e.asText().findText("^(Table|Figure|Image|Diagram|Tabelle|Abbildung|Bild|Diagramm) (\\-?\\d+(\\.\\d+)?)\\. ?");      
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
                alert("Error figureType: "+e);
              };
              if (figureType === null) {
                alert("no figure type! 2");
              };
            };
            // Delete the numbering:
            rangeElement.getElement().asText().deleteText(startOffset,endOffset);
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
      alert('Error in figure: '+error);
    };
    // end of numStyle[eLevel-1] === 'figure'
    // if requested compute new heading numbers
    var txt = '';
    try {      
      if (add == true) {
        numbers[eLevel]++;
        if (eLevel==6 && numStyle[5] === 'figure') {
          // We are at level 6, in a figure.
          // alert("Adding: "+eLevel);
          txt = numbers[1]+"."+numbers[eLevel]+".";
        } else {
          // Build up the number string for chapters
          for (var l = 1; l<=6; l++) {
            if (l <= eLevel) {
              // Note that numStyle[l-1] (0-based) corresponds to numbers[l] (1-based)
              // The heading number:
              var ins = numbers[l];
              // The heading number converted to string (according to style):
              var insX = "";
              // Set the numering style based on numStyle, for each heading
              // var numStylel1 = numStyle[l-1];
              txt = txt + getNumberingStyle(l, numbers[l], numStyle, substr );   
              if (eLevel==l-1 && numStyle[l-1] == 'figure') {
                alert('Figure label now: '+txt);
                txt = getNumberingStyle(l, numbers[l], numStyle, substr );                 
              };
            } else {
              // i.e. l > eLevel
              if ((eLevel!=1 && l==6 && numStyle[5] == 'figure') || numStylel1 === 'T1abcT2' || numStylel1 === 'figure') {
              } else {
                numbers[l] = 0;
              };
            }
            // prefixchar, postfixchar
            var prefixch = "";
            if (prefixchar && prefixchar[eLevel-1]) {
              prefixch = prefixchar[eLevel-1] ;
            }
            txt = prefixch + txt;
            var postfixch = "";
            if (postfixchar && postfixchar[eLevel-1]) {
              postfixch = postfixchar[eLevel-1] ;
            }
            txt = txt + postfixch;
          }
        };
        // done building string
        Logger.log(eText);
        // record value in hcdict, so it can be replaced inreferences
        if (eLevel <= maxLevel || (numStyle[eLevel-1] === 'figure')) {
          var prefixtext = "";
          if (prefixstr && prefixstr[eLevel-1]) {
            prefixtext = prefixstr[eLevel-1] ;
            // alert(eLevel+", t="+prefixtext);
          } else {
            // alert(eLevel+"oooo, t="+prefixtext);
          };
          // collect the new numbering before perfix is added:
          // if (eLevel <= maxLevel ) @+...
          // if (numStyle[eLevel-1] === 'figure') #+...
          if (numStyle[eLevel-1] === 'figure') {
            fndict[currhn] = txt;
            fndict[currhn]  = fndict[currhn].replace(/\. *$/,"");
            // alert("fndict: currhn="+currhn+"->"+txt);
          } else {
            hndict[currhn] = txt;
            hndict[currhn]  = hndict[currhn].replace(/\. *$/,"");
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
          if (eLevel <= maxLevel || (numStyle[eLevel-1] === 'figure')) {
            var style = e.getAttributes();
            var text = e.getText();
            var placeholderStart = 0;
            var placeholderEnd = 0;
            var parent = e.getParent(); 
            var d = DocumentApp.getActiveDocument();
            var parPosition = parent.getChildIndex(e);
            var newPara = d.insertParagraph(parPosition, txt+' ');
            newPara.setAttributes(style);
            try {
              e.merge(); // merge these two paragraphs       
            } catch (e) {
              // collect error messages, and show at the end
              errors += e + "\n";
            }
          } else {
            Logger.log("Not replaced, maxLevel="+maxLevel);
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
    DocumentApp.getUi().alert('There were errors.\n'+errors);
  }
  try {
    if (changeBodyRefs) {
      var style = {};
      style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000'; // null
      style[DocumentApp.Attribute.BACKGROUND_COLOR] = null; // null
      style[DocumentApp.Attribute.FONT_SIZE] = 11;
      regexpRestyleOffset("@\\d",style,0,1);
      regexpRestyleOffset("#\\d",style,0,1);  
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
      for(var key in hndict) {
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
      for(var key in fndict) {
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
      DocumentApp.getUi().alert('Chapter numbering\n'+map+'\nFigure numbering\n'+mapf);
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
      for(var key in arr) {    
        if (arr[key] != '' ) {
          var regexp = "@"+arr[key];
          var str = "⁅C"+arr[key]+"⁆";
          //alert(regexp+ " -> " + str);
          singleReplace(regexp, str, false, false);      
        };
      };      
      for(var key in arr) {
        var value = newdict[arr[key]];
        if (value && arr[key]) {
          if (arr[key] != '' && arr[key] != value) {
            // alert('Chapter numbering x\n'+arr[key]+"->"+value);
            //singleReplace("@"+arr[key],"[["+arr[key]+"]][:]"+value,false,false,null);
            try {
              singleReplace("⁅C"+arr[key]+"⁆","@"+value,false,false,null);
            } catch (e) {
              alert("Error @->[:] - "+e);
            };
          };
        };
      }; 
      var farr = Object.keys(fnewdict);
      for(var key in farr) {      
        if (farr[key] != '' ) {
          var regexp = "#"+farr[key];
          var str = "⁅F"+farr[key]+"⁆";
          //alert("Replace: "+regexp+ " -> "+str);
          singleReplace(regexp, str, false, false);      
        };
      };
      for(var key in farr) {
        var value = fnewdict[farr[key]];
        if (value && farr[key]) {
          if (farr[key] != '' && farr[key] != value) {
            // alert('Chapter numbering x\n'+arr[key]+"->"+value);
            //singleReplace("@"+arr[key],"[["+arr[key]+"]][:]"+value,false,false,null);
            try {
              //alert("Replace: F"+farr[key]+ " -> #"+value);
              singleReplace("⁅F"+farr[key]+"⁆","#"+value,false,false,null);
            } catch (e) {
              alert("Error #->[:] - "+e);
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
        singleReplace("[:]","@",false,false,null);
        // reinstate references
        /* singleReplace("⁅C","@", false, false);      
        singleReplace("⁅F","#", false, false);      
        singleReplace("⁆","", false, false);            */
      } catch (e) {
        alert("Error [:]->@ - "+e);
      };
    }
  } catch (e) {
    alert("Error in changeBodyRefs: "+e);
  };
  // end
}

function getNumberingStyle(l, ins, numStyle, substr) {
  var figureType = "Figure";
  numStylel1 = numStyle[l-1];
  var insX = "";
  var txt = "";
  if (numStylel1 === 'A') {
    insX = String.fromCharCode(64+ins);
  } else if (numStylel1 === '1') {
    insX = ins.toString();
  } else if (numStylel1 === 'S1') {
    insX = "S"+ins.toString();
  } else if (numStylel1 === '§1') {
    insX = "§"+ins.toString();
  } else if (numStylel1 === 'A.1') {
    insX = "A."+ins.toString();
  } else if (numStylel1 === 'B.1') {
    insX = "B."+ins.toString();
  } else if (numStylel1 === 'C.1') {
    insX = "C."+ins.toString();
  } else if (numStylel1 === 'A1') {
    insX = "A"+ins.toString();
  } else if (numStylel1 === 'B1') {
    insX = "B"+ins.toString();
  } else if (numStylel1 === 'C1') {
    insX = "C"+ins.toString();
  } else if (numStylel1 === 'A-1') {
    insX = "A-"+ins.toString();
  } else if (numStylel1 === 'B-1') {
    insX = "B-"+ins.toString();
  } else if (numStylel1 === 'C-1') {
    insX = "C-"+ins.toString();
  } else if (numStylel1 === 'A1-1') {
    insX = "A1-"+ins.toString();
  } else if (numStylel1 === 'B1-1') {
    insX = "B1-"+ins.toString();
  } else if (numStylel1 === 'B2-1') {
    insX = "B2-"+ins.toString();
  } else if (numStylel1 === 'C1-1') {
    insX = "C"+ins.toString();
  } else if (numStylel1 === 'B1-') {
    insX = "B"+ins.toString();
  } else if (numStylel1 === 'B1-') {
    insX = "B"+ins.toString();
  } else if (numStylel1 === '§1-') {
    insX = "§"+ins.toString();
  } else if (numStylel1 === '[1]') {
    insX = "["+ins.toString()+"]";     
  } else if (numStylel1 === 'Teil1') {
    insX = ins.toString();
    if (eLevel == 1) {
      if (isAppendix) {
        insX = "APPENDIX~"+ins.toString();
      } else {
        insX = "KAPITEL~"+ins.toString();
      };
    };
  } else if (numStylel1 === 'T1abcT2') {
    // This style has an odd numbering - used in DFID proposal.
    ins--;
    switch(ins) {
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
        var insm = ins-2;
        insX = "T"+insm.toString()
    }
  } else if (numStylel1 === 'figure') {
    if (figureType == null) {
      alert("no figure type! 10");
      figureType = "Figure";
    };
    // insX = "⸢"+figureType+ " " + ins.toString() + ".⸥"; 
    insX = ""+figureType+ " " + ins.toString() + "."; 
  } else if (numStylel1 === 'PUP') {
    insX = substr[l-1];
  } else {
    insX = "X"+ ins.toString();
  };          
  // Special treatment of 'A1-' styles
  if (l-2 >= 0) {
    switch(numStyle[l-2]) {
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
  switch(numStylel1) {
    case '§1-':
      txt += insX +'-';
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
      txt += insX +'.';
  }
  return txt;
}

function updateSectionNumbersT(){
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var p = doc.getParagraphs();
  var numbers = 0;
  var TS = '';
  for (var i in p) {
    var e = p[i];
    var eText = e.getText()+'';
    var eTypeString = e.getHeading()+'';
    if (!eTypeString.match(/Heading 6/)) {
      continue;
    }

    numbers++;
    if (TS == '') {
      TS = (eText.match(/^(T\d+\-\d+)/))[0];
    };
    Logger.log(eText);
    var newText = eText.replace(/^(T\d+\-\d+)/, TS);
    var newText = newText.replace(/ (\#1|\d+) /, ' '+numbers+ ' ');
    e.setText(newText);
    Logger.log([newText]);
  }
}


function prefixHeadings(){
  var prefix = getValueFromUser("All headings will be prefixed with a fixed string. Please enter the string.");
  if (prefix != "") {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var p = doc.getParagraphs();
    for (var i in p) {
      var e = p[i];
      var eText = e.getText()+'';
      var eTypeString = e.getHeading()+'';
      if (!eTypeString.match(/Heading \d/)) {
        // continue if the paragraph is not a heading
        continue;
      }
//      if (eTypeString.match(/Heading 6/)) {        continue;      }
//      if (eTypeString.match(/Heading 5/)) {        continue;      }
//      if (eTypeString.match(/Heading 4/)) {        continue;      }
      e.editAsText().insertText(0,prefix);
    }
  };
}

/*
if (
        (e.getParent().getType().toString() == 'BODY_SECTION') 
        && (e.getType().toString() == 'PARAGRAPH' || e.getType().toString().match(/Heading \d/))
        ) {          
          number++;
          try {
            var parent = e.getParent(); 
            //DocumentApp.getUi().alert(e.getType());
            var parPosition = parent.getChildIndex(e);
            var newPara = doc.insertParagraph(parPosition, txt);
            newPara.setAttributes(style);
            var thisElementText = newPara.editAsText();
            if (style == 0) {
            if (myfontsize) thisElementText.setFontSize(myfontsize);
              if (fgcolor) thisElementText.setForegroundColor(fgcolor);
              if (bgcolor) thisElementText.setBackgroundColor(bgcolor);
            } else { // if (style==1) {
              thisElementText.setBold(true);
            };
            e.merge(); // merge these two paragraphs       
          } catch (exep) {
            // These should be recorded somehow.
            // DocumentApp.getUi().alert('exception: '+e.getType() + "; " + exep);
            // DocumentApp.getUi().alert("Sorry, this feature hasn't been implemented yet.");
          }
        } else if (
          (e.getParent().getType().toString() == 'TABLE_CELL') 
          && (e.getType().toString() == 'PARAGRAPH' || e.getType().toString().match(/Heading \d/))
        ) {
          // We're in a table
          number++;
          try {
            var parent = e.getParent(); 
            // parent[TABLE_CELL] > e[PARAGRAPH]
            var parPosition = parent.getChildIndex(e);
            // DocumentApp.getUi().alert("Location:"+e.getParent().getType().toString()+">"+e.getType().toString()+">"+parPosition);
            var newPara = parent.insertParagraph(parPosition, txt);
            newPara.setAttributes(style);
            var thisElementText = newPara.editAsText();
            if (style == 0) {
              if (myfontsize) thisElementText.setFontSize(myfontsize);
              if (fgcolor) thisElementText.setForegroundColor(fgcolor);
              if (bgcolor) thisElementText.setBackgroundColor(bgcolor);
            } else { // if (style==1) {
              thisElementText.setBold(true);
            };
            e.merge(); // merge these two paragraphs       
          } catch (exep) {
            // These should be recorded somehow.
            // DocumentApp.getUi().alert('exception: '+e.getType() + "; " + exep);
            // DocumentApp.getUi().alert("Sorry, this feature hasn't been implemented yet.");
          }
        } else if (e.getParent().getType().toString() == 'BODY_SECTION' && e.getType().toString() == 'LIST_ITEM') {
          number++;
          // Cannot merge list items: Element must be preceded by an element of the same type.
          try {
            var parent = e.getParent(); 
            //DocumentApp.getUi().alert(e.getType());
            var parPosition = parent.getChildIndex(e);
            var newPara = doc.insertListItem(parPosition, txt);
            newPara.setAttributes(style);
            var thisElementText = newPara.editAsText();
            if (style == 0) {
              if (myfontsize) thisElementText.setFontSize(myfontsize);
              if (fgcolor) thisElementText.setForegroundColor(fgcolor);
              if (bgcolor) thisElementText.setBackgroundColor(bgcolor);
            } else { // if (style==1) {
              thisElementText.setBold(true);
            };
            // This doesn't work
            // e.merge(); // merge these two paragraphs       
            // but this does:
            var pp2 = parent.getChildIndex(e);
            // var pp1 = parent.getChildIndex(newPara);
            //DocumentApp.getUi().alert(parPosition + ", " + pp1 + "; " + pp2);
            doc.getChild(pp2).merge();
          } catch (exep) {
            // DocumentApp.getUi().alert(e.getType() + "; " + newPara.getType());
            // DocumentApp.getUi().alert('exception: '+e.getType() + "; " + exep);
          }
        } else {
          // e.g. tables are skipped.
          // DocumentApp.getUi().alert('Unknown type: '+e.getType());
        }
        */

  
  
  