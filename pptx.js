
var fs = require('fs');
var unzip = require("unzip");
var DOMParser = require('xmldom').DOMParser;

module.exports.readPptxFile = readPptxFile;

function readPptxFile(filename, entryMatchFunc, docHandler, callback) {

  var unzipParser = unzip.Parse();
  var result = [];

  fs.createReadStream(filename)
    .pipe(unzipParser)
    .on('entry', function (entry) {
      // for each file, check if the caller is interested.
      if (entryMatchFunc(entry)) {
        // OK, create a scope and handle the data from the stream
        (function (entry) {
          var xml = '';
          // assume we have utf-8 encoding
          entry.setEncoding('utf8');
          entry.on('data', function (chunk) {
            // some data has arrived from the stream, append it to the XML-buffer
            xml += chunk;
          });
          entry.on('end', function () {
            // no more data for this entry, parse the XML and call the document handler function.
            var data = docHandler( new DOMParser().parseFromString( xml ) );
            if (data) {
              // if the handler could handle the document, push the data to the result.
              result.push(data);
            }
          });
        }) (entry);
      } else {
        // this entry isn't interesting, skip it.
        entry.autodrain();
      }
    });

  unzipParser.on('close', function () {
    // the file has been processed, just send the result back
    if (callback) {
      callback(result);
    }
  });
}
