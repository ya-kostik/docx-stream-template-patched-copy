const Template = require('stream-text-variable-template');
const unzip = require('unzip');
const zip = require('zip-stream');

function docxTemplate(props, docx) {
  const archive = new zip({ zlib: { level: 0 } });
  let promise = Promise.resolve();
  function addFile(archive, stream, fileName) {
    return new Promise((resolve, reject) => {
      setImmediate(() => {
        archive.entry(stream, { name: fileName }, function(err) {
          if (err) return reject(err);
          process.nextTick(() => resolve());
        });
      });
    });
  }
  docx.pipe(unzip.Parse()).
  on('entry', function (entry) {
    var fileName = entry.path;
    if (fileName === "word/document.xml") {
      const stream = entry.pipe(new Template(props));
      promise = promise.then(() => addFile(archive, stream, fileName));
    } else {
      if (fileName === '[Content_Types].xml') {
        setImmediate(() => {
          promise = promise.then(() => {
            return addFile(archive, entry, fileName);
          }).
          then(() => archive.finalize());
          // archive.entry(entry, { name: fileName })
          // archive.finalize();
        });
      } else {
        promise = promise.then(() => addFile(archive, entry, fileName));
        // setImmediate(() => archive.entry(entry, { name: fileName }));
      }
    }
  });
  archive.once('error', () => {
    docx.destroy();
  });
  return archive;
}

module.exports = docxTemplate;
