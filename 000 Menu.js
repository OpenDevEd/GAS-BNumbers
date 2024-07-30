function onOpen() {
  Menu_HePaNumbering().addToUi();
}

function Menu_HePaNumbering() {
  const activeStyle = getHeadingStyle();
  const debugMode = getDebugMode();

  // Select H1 prefix
  const subMenuPrefixes = DocumentApp.getUi().createMenu('Select prefix');
  let selectedPrefixMarker = '';
  for (let prefix in prefixes) {
    if (prefix.toString() == activeStyle.prefix) {
      selectedPrefixMarker = activeStyle.prefixMarker;
    } else {
      selectedPrefixMarker = '';
    }
    subMenuPrefixes.addItem(prefixes[prefix]['name'] + ' ' + selectedPrefixMarker, 'prefixes.' + prefix + '.run');
  }
  // End. Select H1 prefix

  // Select style submenu
  const subMenuStyles = DocumentApp.getUi().createMenu('Select style')
    .addItem('Show current style', 'showCurrentStyle')
    .addSubMenu(DocumentApp.getUi().createMenu('Configure ignored text')
      .addItem('Show current setting', 'ignoredTextShowCurrentSetting')
      .addItem('Edit setting', 'ignoredTextEditSetting')
      .addItem('Restore default', 'ignoredTextRestoreDefault'));

  let selectedStyleMarker = '';
  let menuItemText;
  for (let styleName in headingStyles) {
    menuItemText = headingStyles[styleName]['name'];
    if (styleName.toString() === 'boldFugureH6') {
      const boldFugureH6Style = getBoldFigureStyle(true);
      selectedStyleMarker = boldFugureH6Style.marker;
    } else {
      if (styleName.toString() == activeStyle.value) {
        selectedStyleMarker = activeStyle.marker;
      } else {
        selectedStyleMarker = '';
      }
    }
    if (headingStyles[styleName]['separatorAbove'] === true) {
      subMenuStyles.addSeparator();
    }
    if (headingStyles[styleName]['overrideH1PrefixWithCustomPrefix'] === true || headingStyles[styleName]['overrideAllHPrefixWithCustomPrefix'] === true) {
      menuItemText = menuItemText.replaceAll('~PREFIX~', activeStyle.prefixText);
    }
    subMenuStyles.addItem(menuItemText + ' ' + selectedStyleMarker, 'headingStyles.' + styleName + '.run');
    if (headingStyles[styleName]['subMenuBelow'] === true) {
      subMenuStyles.addSubMenu(subMenuPrefixes);
    }
  }
  // End. Select style submenu

  const submenu_util = DocumentApp.getUi().createMenu('Utilities')
    .addItem('Add/update heading numbers only [links not updated]', 'doNumberHeadings')
    .addSeparator()
    .addItem('Links: Update numbers in (hyper)links (to headings)', 'updateNumbersInLinksToHeadings')
    .addItem('Links: Replace entire link text with text of corresponding heading', 'resetFullLinkTextInLinksToHeadings')
    .addSeparator()
    .addItem('nhr Remove all heading numbers', 'removeAllHeadingNumbers')
    .addSeparator()
    .addItem('Mark internal heading links', 'markInternalHeadingLinks')
    .addItem('Clear internal heading link markers', 'clearInternalLinkMarkers')
    .addSeparator()
    .addItem('nhprefix prefix all headings with a string', 'prefixHeadings')
    .addItem('nh6t Heading 6 tables/figures only', 'updateFigureNumbers');

  const submenu_para = DocumentApp.getUi().createMenu('Paragraphs')
    .addItem('pna Paragraph numbers add, with  ⟦ and ⟧', 'paraNumAdd')
    .addItem('psna Paragraph/sentence numbers add, with ⟦ and ⟧', 'paraSenNumAdd')
    .addItem('psnr Paragraph/sentence numbers remove', 'paraSenNumRemove')
    .addSeparator()
    .addItem('psnm Paragraph/sentence numbers minify', 'minifyParaSenMarker')
    .addItem('psnshow Paragraph/sentence numbers - restore size', 'maxifyParaSenMarker');

  const menu = DocumentApp.getUi().createMenu('Heading & paragraph numbering')
    .addItem('nha Add/update heading numbers and links', 'doNumberHeadingsAndLinks')
    .addSeparator()
    .addSubMenu(subMenuStyles)
    .addSubMenu(submenu_util)
    .addSubMenu(submenu_para);

  if (debugMode === true) {
    menu.addItem('Turn off debug mode', 'turnOffDebugMode');
  } else {
    menu.addItem('Turn on debug mode', 'turnOnDebugMode');
  }

  return menu;
}