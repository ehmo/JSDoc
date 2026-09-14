// SPDX-License-Identifier: 0BSD
/*
 * writer.test.js — validate textToDoc() by round-tripping the .doc it writes.
 *
 * This cross-platform test reads the output back with two independent parsers:
 * our own docToText (must reproduce the input exactly) and the unrelated
 * word-extractor (must parse it and find the text). Both succeeding is strong
 * evidence the binary is well-formed. The docToText round-trip is offline and
 * always runs; the word-extractor check skips if that dev dep isn't installed.
 *
 * Parsers are lenient, though — that's exactly why a from-scratch .doc looked
 * fine here but failed in a real editor. The decisive check is opening the
 * output in an actual word processor; see scripts/read-with-textmaker.ps1,
 * which drives SoftMaker TextMaker via COM (Windows-only, hence not in CI).
 */
'use strict';
var textToDoc = require('../src/textToDoc.js');
var docToText = require('../src/docToText.js');

var failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'ok  ' : 'FAIL') + ' ' + name);
  if (!cond) failures++;
}

console.log('textToDoc — write a .doc and read it back\n');

var cases = [
  'Hello, World!',
  'Line one\nLine two\nLine three',
  'Smart quotes “like this”, an em dash —, and accents éàü.',
  'Unicode: π ∑ ✓ 😀',
  ''
];

cases.forEach(function (input, i) {
  var bytes;
  try { bytes = textToDoc(input); } catch (e) { check('case ' + i + ' writes', false); return; }
  check('case ' + i + ' produces bytes', bytes && bytes.length >= 512 && bytes[0] === 0xD0 && bytes[1] === 0xCF);
  var back = docToText(bytes);
  // textToDoc appends a trailing paragraph mark; compare on normalized lines.
  var want = input.replace(/\r\n?|\n/g, '\n');
  var got = (back || '').replace(/\r\n?|\n/g, '\n').replace(/\n$/, '');
  check('case ' + i + ' round-trips through docToText', got === want);
});

// The bundled template is a Word 2002 FIB. Its output must declare the
// matching FibRgCswNew extension rather than relying on lenient readers.
(function () {
  var streams = textToDoc.readCfb(textToDoc('FIB profile check.'));
  var wordDocument = streams && streams.byName.WordDocument;
  var view = wordDocument && new DataView(wordDocument.buffer,
    wordDocument.byteOffset, wordDocument.byteLength);
  var csw = view && view.getUint16(32, true);
  var cslwAt = 34 + csw * 2;
  var cslw = view && view.getUint16(cslwAt, true);
  var pairCountAt = cslwAt + 2 + cslw * 4;
  var pairCount = view && view.getUint16(pairCountAt, true);
  var cswNewAt = pairCountAt + 2 + pairCount * 8;
  check('Word 2002 FIB has 136 FibRgFcLcb pairs', pairCount === 0x0088);
  check('Word 2002 FIB declares two FibRgCswNew words', view.getUint16(cswNewAt, true) === 2);
  check('Word 2002 FIB extension repeats nFib 0x0101', view.getUint16(cswNewAt + 2, true) === 0x0101);
  check('Word 2002 FIB extension reserved word is zero', view.getUint16(cswNewAt + 4, true) === 0);
})();

// CFB files with more than 109 FAT sectors need a DIFAT continuation sector.
// This covers both the writer and its exposed template/tooling reader.
var large = new Uint8Array(8 * 1024 * 1024);
for (var lp = 0; lp < large.length; lp += 4093) large[lp] = lp / 4093 & 0xFF;
var largeCfb = textToDoc.buildCfb([{ name: 'Large', data: large }]);
var largeView = new DataView(largeCfb.buffer, largeCfb.byteOffset, largeCfb.byteLength);
var largeRead = textToDoc.readCfb(largeCfb);
check('large CFB uses more than 109 FAT sectors', largeView.getUint32(44, true) > 109);
check('large CFB writes a DIFAT continuation sector', largeView.getUint32(72, true) === 1);
check('large CFB reader follows DIFAT', !!largeRead && !!largeRead.byName.Large
  && largeRead.byName.Large.length === large.length
  && largeRead.byName.Large[0] === large[0]
  && largeRead.byName.Large[4093 * 1024] === large[4093 * 1024]);

// Independent oracle: word-extractor must also parse our .doc.
(function () {
  var WordExtractor;
  try { WordExtractor = require('word-extractor'); }
  catch (e) { console.log('\n  skip word-extractor cross-check (not installed)'); return done(); }
  var input = 'Independent reader check.\nSecond paragraph.';
  var buf = Buffer.from(textToDoc(input));
  new WordExtractor().extract(buf).then(function (doc) {
    var body = doc.getBody();
    check('word-extractor parses our .doc and finds the text',
      body.indexOf('Independent reader check.') !== -1 && body.indexOf('Second paragraph.') !== -1);
    done();
  }).catch(function (e) {
    check('word-extractor parses our .doc (' + e.message + ')', false);
    done();
  });
})();

function done() {
  console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)');
  process.exit(failures === 0 ? 0 : 1);
}
