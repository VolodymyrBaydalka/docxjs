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
    'equation',
    'text-box',
    'text-box-wps',
    'image',
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

      if(actual != expected) {
        const mismatch = findFirstMismatch(expected, actual);

        console.log(`[-] ${JSON.stringify(mismatch.expected)}`);
        console.log(`[+] ${JSON.stringify(mismatch.actual)}`);

      }

      expect(actual).toBe(expected);

      div.remove();
    });
  }
});

function formatHTML(text) {
  return text
    .replace(/src="blob:[^"]+"/ig, 'src="blob:__dynamic__"')
    .replace(/\t+|\s+/ig, ' ')
    .replace(/<style>\s+/ig, '<style>')
    .replace(/\s+<\/style>/ig, '</style>')
    .replace(/\{\}/ig, '{ }')
    .replace(/>\s+</ig, '><')
    .replace(/></ig, '>\n<')
    .trim();
}

function findFirstMismatch(expected, actual) {
  const max = Math.min(expected.length, actual.length);

  for(let i = 0; i < max; i++) {
    if(expected[i] !== actual[i]) {
      return {
        index: i,
        expected: snippetAt(expected, i),
        actual: snippetAt(actual, i),
      };
    }
  }

  return {
    index: max,
    expected: snippetAt(expected, max),
    actual: snippetAt(actual, max),
  };
}

function snippetAt(text, index, radius = 80) {
  const start = Math.max(0, index - radius);
  const end = Math.min(text.length, index + radius);
  return text.slice(start, end);
}
