
var pptx = require('./pptx');

module.exports.extractNotes = extractNotes;

function extractNotes(filename, callback) {

  // start the pptxReading.
  pptx.readPptxFile(filename, entryMatchFunc, docHandler, resultHandler);

  function entryMatchFunc(entry) {
    // notes are located in ppt/notesSlides/noteSlideX.xml
    return entry.path.match(/^ppt\/notesSlides\/notesSlide\d+\.xml/);
  }

  function docHandler(doc) {
    var txt, fldElm, rows;
    var note = '';
    var slideNumber;

    function rowHandler(row) {
      // find all text-elements (<a:t>)
      var elements = row.getElementsByTagName('a:t');
      var textBuffer = '';
      for (var i = 0; i < elements.length; i++) {
        if (elements[i].textContent) {
          // add the textContent to the text buffer and an extra space at the end.
          textBuffer += elements[i].textContent.trim() + " ";
        }
      }
      // remove extra spaces and return text-buffer
      return textBuffer.trim();
    }

    // Find all textRows (<a:p>)
    rows = doc.getElementsByTagName('a:p');
    for (var i = 0; i < rows.length; i++) {
      // extract the text on this row
      txt = rowHandler(rows[i]);
      // check if this row is a slide-number (<a:p><a:fld type="slidenum">...</a:fld></a:p>)
      fldElm = rows[i].getElementsByTagName('a:fld');
      if (fldElm.length && fldElm[0].getAttribute('type') == 'slidenum') {
        // it is a slide-number
        slideNumber = parseInt(txt, 10);
      } else {
        // not a slide-number, add the text to the note data and add a newline
        note += txt + "\n";
      }
    }
    // remove unwanted whitespace
    note = note.trim();
    if (note === '') {
      // if there is no data at all, return undefined instead of empty string
      return undefined;
    }

    return {
      'slideNumber' : slideNumber,
      'note' : note
    };
  }

  function resultHandler(result) {
    // sort the notes in slide-number order.
    result.sort(function (a,b) {
      return a.slideNumber - b.slideNumber;
    });
    if (callback) {
      callback(result);
    }
  }
}