
function paraNumAdd() {
  sequentialNumbers(true, false, 0);
}

function paraNumRemove() {
  sequentialNumbers(false, false, 0);
}

function paraNumAddStyle1() {
  sequentialNumbers(true, false, 1);
}


function paraSenNumAddPlus() {
  // Get prefix from user.
  // sequentialNumbers(true,true,-1);
  //sequentialNumbers(true,true,0);
  //sequentialNumbers(true,true,1);
  sequentialNumbersX(true, true, 2, 1);
}

function paraSenNumAdd() {
  sequentialNumbersX(true, true, 2, 0);
}

function paraSenNumRemove() {
  sequentialNumbers(false, true, 10);
}

function minifyParaSenMarker() {
  var style = {};
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FEFEFE'; // null
  //style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#AAAAAA'; // null
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  style[DocumentApp.Attribute.FONT_SIZE] = 1;
  //  singleReplace("§6.","§⍶6.",false,false,null);
  //  singleReplace("§⍶ 6.","§⍶6.",false,false,null);
  singleReplace("⌂ ", "⌂", false, false, null);
  regexpRestyle("⌂", style);
  //  singleReplace("⇡(","(⇡",false,false,null);
  //  style = {};
  //  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#103DC1'; // null
  //  regexpRestyle("⇡",style) ;
  //  regexpRestyle("\\\(⇡",style) ;
  //  singleReplace("⸣","",false,false,null);
  //  singleReplace("ibid.","ebd.",false,false,null);
}