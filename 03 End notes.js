function moveParagraphNotes_afterRefs_tablesAndLists_single() {
  moveParagraphNotes_afterRefs_tablesAndLists(false);
}

function moveParagraphNotes_afterRefs_tablesAndLists_multiple() {
  moveParagraphNotes_afterRefs_tablesAndLists(true);
}

function notesExistenceCheck(body) {
  var DIVIDER_RE_DOC = "^[ \t\n\r\f\v]*⟦\/Notes\/⟧[ \t\n\r\f\v]*";
  const dividerEl = body.findText(DIVIDER_RE_DOC);
  if (dividerEl == null) {
    alert("Please place the notes at the end of your document, using ⟦/Notes/⟧ to indicate the start of the notes.");
    return { status: false };
  } else {
    return { status: true, dividerEl};
  }
}

function moveParagraphNotes_afterRefs_tablesAndLists(multiple) {
  var body = DocumentApp.getActiveDocument().getBody();

  var lastIndex = body.getNumChildren() - 1;

  // Regex
  var DIVIDER_RE = /^\s*⟦\/Notes\/⟧\s*$/;
  var MARKER_RE = /^\s*⟦(\d+)⟧/;
  var NOTE_RE = /^\s*⟦(\d+)⟧\s*.*$/;

  // const dividerFlag = body.findText(DIVIDER_RE_DOC);
  // if (dividerFlag == null) {
  //   alert("Please place the notes at the end of your document, using ⟦/Notes/⟧ to indicate the start of the notes.");
  //   return 0;
  // }
  const { status } = notesExistenceCheck(body);
  if (!status) return 0;

  // 1) Traverse in reading order; collect paragraphs & list items, split by divider
  var before = []; // elements before divider (Paragraph or ListItem)
  var after = []; // elements after divider  (Paragraph or ListItem)
  var state = { after: false };

  walkContainer_(body, state, before, after, DIVIDER_RE);

  // after.forEach(para => Logger.log(para.getText()));

  // 2) Build map of targets (unique per spec): num -> element
  var targetsByNum = {};
  var targetsOrder = []; // {num, el, order} to process bottom-up
  for (var i = 0; i < before.length; i++) {
    var el = before[i];
    var txt = getElementText_(el);
    var m = txt ? txt.match(MARKER_RE) : null;
    if (m) {
      var num = m[1];
      targetsByNum[num] = el;
      targetsOrder.push({ num: num, el: el, order: i });
    }
  }

  // 3) Collect notes in 'after' grouped by number (keep document order)
  var notesByNum = {}; // { "1": [Element, ...] }
  let nnum;
  for (var j = 0; j < after.length; j++) {
    var nel = after[j];
    var ntext = getElementText_(nel);
    var nm = ntext ? ntext.match(NOTE_RE) : null;
    if (!nm && multiple === false || ntext.length === 0) continue;
    if (nm) {
      nnum = nm[1];
      if (!notesByNum[nnum]) notesByNum[nnum] = [];
    }
    notesByNum[nnum].push(nel);
  }

  // 4) Process referenced elements bottom-up (stable indices)
  targetsOrder.sort(function (a, b) { return b.order - a.order; });

  for (var t = 0; t < targetsOrder.length; t++) {
    var numKey = targetsOrder[t].num;
    var refEl = targetsOrder[t].el;
    var noteEls = notesByNum[numKey];
    if (!noteEls || !noteEls.length) continue;

    // Find the direct container (Body or TableCell) and index of the referenced element
    var refContainer = directContainer_(refEl);
    if (!refContainer) continue;
    var refIndex = refContainer.getChildIndex(refEl);

    // Insert each note in reverse order right after the reference
    for (var r = noteEls.length - 1; r >= 0; r--) {
      var noteEl = noteEls[r];

      // Clone the note element (Paragraph or ListItem) to preserve formatting
      var clone = noteEl.copy();

      // Convert prefix "⟦n⟧ ..." into "⟦:n:...⟧"
      convertNotePrefixGeneric_(clone, numKey);

      // Insert after the reference in the same container, preserving element type
      if (clone.getType() === DocumentApp.ElementType.PARAGRAPH) {
        refContainer.insertParagraph(refIndex + 1, clone.asParagraph());
      } else if (clone.getType() === DocumentApp.ElementType.LIST_ITEM) {
        refContainer.insertListItem(refIndex + 1, clone.asListItem());
      } else {
        // Fallback: insert as paragraph
        var para = DocumentApp.createParagraph(getElementText_(clone));
        refContainer.insertParagraph(refIndex + 1, para);
      }

      // Remove original note from wherever it lives (Body or TableCell)
      var noteContainer = directContainer_(noteEl);
      if (noteContainer) {
        if (noteContainer.getType() === DocumentApp.ElementType.BODY_SECTION && lastIndex === noteContainer.getChildIndex(noteEl)) {
          // noteEl.setText(" ");
          noteEl.clear();
        } else {
          noteContainer.removeChild(noteEl);
        }
      }
    }
  }

  // Unmatched notes remain in place under ⟦/Notes/⟧.
}

function elementsEqual(element1, element2) {
  // if (!element1 || !element2) return false;

  var parent1 = element1.getParent();
  var parent2 = element2.getParent();

  // Check if same parent and same index
  return parent1.getType() === parent2.getType() &&
    parent1.getChildIndex(element1) === parent2.getChildIndex(element2);
}

/** Walk Body/Cells recursively in reading order; collect Paragraphs and ListItems; split by divider. */
function walkContainer_(container, state, before, after, DIVIDER_RE) {
  var n = container.getNumChildren();
  for (var i = 0; i < n; i++) {
    var ch = container.getChild(i);
    var type = ch.getType();

    if (type === DocumentApp.ElementType.TABLE) {
      var table = ch.asTable();
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        for (var c = 0; c < row.getNumCells(); c++) {
          walkContainer_(row.getCell(c), state, before, after, DIVIDER_RE);
        }
      }
    } else if (type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM) {
      var el = (type === DocumentApp.ElementType.PARAGRAPH) ? ch.asParagraph() : ch.asListItem();
      var txt = getElementText_(el) || '';
      if (DIVIDER_RE.test(txt)) {
        state.after = true; // everything after this in reading order goes to 'after'
        continue;           // we don't need to store the divider itself
      }
      (state.after ? after : before).push(el);
    } else {
      // Ignore other element types for this task
    }
  }
}

/** Get the direct container that supports insert/remove child: Body or TableCell. */
function directContainer_(el) {
  var p = el.getParent();
  while (p &&
    p.getType() !== DocumentApp.ElementType.BODY_SECTION &&
    p.getType() !== DocumentApp.ElementType.TABLE_CELL) {
    p = p.getParent();
  }
  return p || null;
}

/** Get plain text of a Paragraph or ListItem. */
function getElementText_(el) {
  // Both Paragraph and ListItem support getText()
  if (el.getType && (el.getType() === DocumentApp.ElementType.PARAGRAPH ||
    el.getType() === DocumentApp.ElementType.LIST_ITEM)) {
    return el.getText();
  }
  // Fallback for safety
  if (typeof el.getText === 'function') return el.getText();
  return '';
}

/**
 * Transform a cloned note element (Paragraph or ListItem) from:
 *   "⟦n⟧   <note text...>"  →  "⟦:n:<note text...>⟧"
 * Keeps all inline styles by only deleting the prefix across text nodes,
 * then inserting small wrapper bits.
 */
function convertNotePrefixGeneric_(el, num) {
  // We need Paragraph/ListItem API (children of type TEXT)
  var full = getElementText_(el) || '';
  var prefixRe = new RegExp('^\\s*⟦' + num + '⟧\\s*');
  var m = full.match(prefixRe);
  var removeLen = m ? m[0].length : 0;

  // Delete the leading prefix across text nodes
  var remaining = removeLen;
  while (remaining > 0) {
    var first = el.getChild && el.getChild(0);
    if (!first || first.getType() !== DocumentApp.ElementType.TEXT) break;
    var t = first.asText();
    var len = t.getText().length;
    if (remaining >= len) {
      el.removeChild(first);
      remaining -= len;
    } else {
      t.deleteText(0, remaining - 1);
      remaining = 0;
    }
  }

  // Prepend '⟦:n:' and append closing '⟧' (small wrapper text with default style)
  if (typeof el.insertText === 'function') {
    el.insertText(0, '⟦:' + num + ':');
  } else if (el.editAsText) {
    el.editAsText().insertText(0, '⟦:' + num + ':');
  }
  if (typeof el.appendText === 'function') {
    el.appendText('⟧');
  } else if (el.editAsText) {
    var et = el.editAsText();
    et.appendText('⟧');
  }
}
