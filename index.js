
var cli = require('cli').enable('version', 'status'), options = cli.parse({
  input: ['i', 'path to .pptx file', 'path']
});

cli.main(function (args, options) {

  cli.debug(JSON.stringify(options));

  function readPptxFile(filename, entryMatchFunc, docHandler, callback) {

    var fs = require('fs');
    var unzip = require("unzip");
    var DOMParser = require('xmldom').DOMParser;

    var unzipParser = unzip.Parse();
    var result = [];

    fs.createReadStream(filename)
      .pipe(unzipParser)
      .on('entry', function (entry) {
        if (entryMatchFunc(entry)) {
          (function (entry) {
            cli.debug('Found '+entry.path);
            var xml = '';
            entry.setEncoding('utf8');
            entry.on('data', function (chunk) {
              xml += chunk;
            });
            entry.on('end', function () {
              var data = docHandler( new DOMParser().parseFromString( xml ) );
              if (data) {
                result.push(data);
              }
            });
          }) (entry);
        } else {
          entry.autodrain();
        }
      });

    unzipParser.on('close', function () {
      if (callback) {
        callback(result);
      }
    });
  }


  function extractNotes(filename, callback) {
    cli.debug('Extracting notes from: '+ filename);

    readPptxFile(filename, entryMatchFunc, docHandler, resultHandler);

    function entryMatchFunc(entry) {
      return entry.path.match(/^ppt\/notesSlides\/notesSlide\d+\.xml/);
    }

    function docHandler(doc) {
      function rowHandler(row) {
        var elements = row.getElementsByTagName('a:t');
        var txt = '';
        var isSlideNumber = false;
        for (var i = 0; i < elements.length; i++) {
          if (elements[i].textContent) {
            txt += elements[i].textContent.trim() + " ";
          }
        }
        return txt.trim();
      }

      var rows = doc.getElementsByTagName('a:p');
      var note = '';
      var slideNumber;
      var txt;
      var fldElm;
      for (var i = 0; i < rows.length; i++) {
        txt = rowHandler(rows[i]);
        fldElm = rows[i].getElementsByTagName('a:fld');
        if (fldElm.length && fldElm[0].getAttribute('type') == 'slidenum') {
          slideNumber = parseInt(txt, 10);
        } else {
          note += txt + "\n";
        }
      }
      note = note.trim();
      if (note === '') {
        return null;
      }
      return {
        'slideNumber' : slideNumber,
        'note' : note
      };
    }

    function resultHandler(result) {
      result.sort(function (a,b) {
        return a.slideNumber - b.slideNumber;
      });
      if (callback) {
        callback(result);
      }
    }
  }

  if (options.input) {
    extractNotes(options.input, function (result) {
      console.log(result);
    });
  }

});