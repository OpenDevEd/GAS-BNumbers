function onOpen() {
  submenu_01_HePaNumberingFull().addToUi();
}

function submenu_01_HePaNumberingBasic() {
  return DocumentApp.getUi().createMenu('Heading & paragraph numbering')
    .addItem('nha Add/update heading numbers and links', 'doNumberHeadingsAndLinks')
    .addSubMenu(DocumentApp.getUi().createMenu('Utilities')
      .addItem('Add/update heading numbers only', 'doNumberHeadings')
      .addItem('Update numbers in links to headings', 'updateNumbersInLinksToHeadings')
      // Elena:
      .addItem('Reset full link text in links to headings', 'resetFullLinkTextInLinksToHeadings')
      .addSeparator()
      // Elena:
      .addItem('Mark internal heading links', 'markInternalHeadingLinks')
                /*
                Elena:
                Similar to the broken links with BZotero, show a message: 
                There were broken links. Please search for BROKEN_INTERNAL_LINK_MARKER and fix these. Note that broken links can occur when you press enter at the start of an existing heading./['
                */
      .addItem('Clear internal heading link markers', 'clearInternalLinkMarkers')
      .addSeparator()
      .addItem('Show current style', 'sorryNotImplementedYet')
      .addSeparator()
      .addItem('nhprefix prefix all headings with a string', 'prefixHeadings')
      .addSeparator()
      .addItem('nh6t Heading 6 tables/figures only', 'updateFigureNumbers')
      .addItem('nh6s Heading 6 special: Prefix plus section numbers', 'updateSectionNumbersT')
      .addItem('nh6r remove - not implemented', 'sorryNotImplementedYet')
      .addSeparator()
      .addItem('nhr Remove all heading numbers (H1-6)', 'numberHeadingsClear6')
      //.addItem('nhnr6 Remove heading numbers (H1-6)', 'numberHeadingsClear6')
      .addItem('nhnr4 Remove heading numbers (H1-4)', 'numberHeadingsClear4')
      .addItem('nhnr3 Remove heading numbers (H1-3)', 'numberHeadingsClear3')
    )
    .addSubMenu(DocumentApp.getUi().createMenu('Select style')
      .addItem('Show current style', 'sorryNotImplementedYet')
      .addSeparator()
      .addItem('H1-3/-/-/figure + refs', 'numberHeadingsAdd3Figure')
      .addItem('H1-4/-/figure + refs', 'numberHeadingsAdd4Figure')
      .addItem('H1-5/figure + refs', 'numberHeadingsAddFigure')
      .addItem('Chapter~#/H2-H3/-/figure + refs', 'numberHeadingsAddChapter23')
      .addItem('Chapter~#/H2-H4/-/figure + refs', 'numberHeadingsAddChapter234')
      .addItem('Section~#/H2-H3/-/figure + refs', 'numberHeadingsAddSection23')
      .addItem('Section~#/H2-H4/-/figure + refs', 'numberHeadingsAddSection234')
      .addItem('Session~#/H2-H3/-/figure + refs', 'numberHeadingsAddSession23')
      .addSeparator()
      .addSubMenu(DocumentApp.getUi().createMenu('Select H1 prefix (not implemented)')
        .addItem('<none>', 'sorryNotImplementedYet')
        .addItem('Chapter', 'sorryNotImplementedYet')
        .addItem('Section', 'sorryNotImplementedYet')
        .addItem('Session', 'sorryNotImplementedYet')
        .addItem('<enter>', 'sorryNotImplementedYet')
      )
      .addSeparator()
      .addItem('H1-6 (no refs)', 'numberHeadingsAdd6')
      .addItem('H1-4 (no refs)', 'numberHeadingsAddPartial4')
      .addItem('H1-3 (no refs)', 'numberHeadingsAddPartial3')
      // .addItem('e='+email, 'numberHeadingsAddPartial')
    )
  //.addSubMenu(DocumentApp.getUi().createMenu('Useful')
  //)
};

function submenu_01_HePaNumberingFull() {
  var HePaBasic = submenu_01_HePaNumberingBasic();
  return HePaBasic
    /*
      .addSubMenu(DocumentApp.getUi().createMenu('Additional Styles')
        //.addItem('nhna Add/update heading numbers (Teil 1, 1A11+-+figure)', 'numberHeadingsAddTeilFigure1A')
        //.addItem('nhnb Add/update heading numbers (Teil 1, 1111+-+figure)', 'numberHeadingsAddTeilFigure11')
        .addItem('nhnb Add/update heading numbers (Kapitel~1, 1111+-+figure)', 'numberHeadingsAddTeilFigure11')
        .addItem('nhng Add/update heading numbers (Kapitel~#, 1111+-+figure)', 'numberHeadingsAddTeilFigure11pickGerman')
        .addItem('nhno Add/update heading numbers OER4S', 'numberHeadingsRepeatH1')
        // .addItem('hn4a Add/update heading numbers (1.A.1.1.)', 'numberHeadingsAddPartial1A11')
        .addItem('hnt Add/update heading numbers (B1-1-T1.1.)', 'numberHeadingsAddPartialB11T')
        //.addItem('hntdash Add/update heading numbers (B-1.T1.1.)', 'numberHeadingsAddPartialBdash1T')
        //.addItem('hntdot Add/update heading numbers (B.1.T1.1.)', 'numberHeadingsAddPartialBdot1T')
      )
      */
    .addSubMenu(DocumentApp.getUi().createMenu('Paragraphs')
      .addItem('pna Paragraph numbers add', 'paraNumAdd')
      .addItem('pnr Paragraph numbers remove', 'paraNumRemove')
      .addSeparator()
      .addItem('pna1 Paragraph numbers add, style 1', 'paraNumAddStyle1')
      .addItem('psna Paragraph/sentence numbers add', 'paraSenNumAdd')
      .addItem('psnap Paragraph/sentence numbers - "plus"', 'paraSenNumAddPlus')
      .addSeparator()
      .addItem('psnr Paragraph/sentence numbers remove ', 'paraSenNumRemove')
      .addItem('psnm Paragraph/sentence numbers minify ', 'minifyParaSenMarker')
      .addItem('psnt Paragraph/sentence numberlinks back to text ', 'sequentialNumbersPlusToText')
    )
};

