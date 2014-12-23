
var cli = require('cli').enable('status'), options = cli.parse({
  input: ['i', 'path to .pptx file', 'path']
});

cli.main(function (args, options) {

  if (options.input) {
    require('./pptx_notes').extractNotes(options.input, function (result) {
      console.log(result);
    });
  }

});