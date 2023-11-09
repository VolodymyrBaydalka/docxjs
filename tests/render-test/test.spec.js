describe("Render document", function () {
  const tests = [
    'text',
    'underlines',
    'text-break',
    'table',
    'page-layout',
    'revision',
    'numbering',
    'line-spacing',
    'header-footer',
    'footnote',
    'equation'
  ];

  for (let path of tests) {
    it(`from ${path} should be correct`, async () => {

      const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
      const resultText = await fetch(`/base/tests/render-test/${path}/result.html`).then(r => r.text());

      const div = document.createElement("div");

      document.body.appendChild(div);

      await docx.renderAsync(docBlob, div);
      
      const actual = formatHTML(div.innerHTML);
      const expected = formatHTML(resultText);

      expect(actual).toBe(expected);

      if(actual != expected) {
        const diffs = Diff.diffLines(expected, actual);

        for(const diff of diffs) {
          if(diff.added)
            console.log('[+] ' + diff.value);

          if(diff.removed)
            console.log('[-] ' + diff.value);
        }
      }

      div.remove();
    });
  }
});

function formatHTML(text) {
  return text.replace(/\t+|\s+/ig, ' ').replace(/></ig, '>\n<');
}