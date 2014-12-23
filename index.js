
var cli = require('cli').enable('version', 'status'), options = cli.parse({
  input: ['i', 'path to .pptx file', 'path']
});
var fs = require('fs');
var unzip = require("unzip");
var DOMParser = require('xmldom').DOMParser;


cli.main(function(args, options) {
  cli.debug(JSON.stringify(options));

  var noteData = [];
  if (!options.input) {
    return;
  }

  var unzipParser = unzip.Parse();
  fs.createReadStream(options.input)
    .pipe(unzipParser)
    .on('entry', function (entry) {
      if (entry.path.match(/^ppt\/notesSlides\/notesSlide\d+\.xml/)) {
        (function(entry){
          cli.debug('Found '+entry.path);
          var xml = '';
          entry.setEncoding('utf8');
          entry.on('data', function(chunk) {
            xml += chunk;
          });
          entry.on('end', function() {
            var doc = new DOMParser().parseFromString(xml);
            var elms = doc.getElementsByTagName('a:t');
            var note = '';
            var slideNumber;
            for (var i in elms) {
              if (elms.hasOwnProperty(i) && elms[i].textContent != undefined) {
                if (elms[i].parentNode.getAttribute('type') == 'slidenum') {
                  slideNumber = parseInt(elms[i].textContent, 10);
                } else {
                  note += elms[i].textContent;
                }
              }
            }
            noteData.push({
              'slideNumber' : slideNumber,
              'note' : note
            });
          });
        })(entry);
      } else {
        entry.autodrain();
      }
    });
  unzipParser.on('close', function () {
    noteData.sort(function(a,b){
      return a.slideNumber - b.slideNumber;
    });
    console.log(noteData);
  });

});