describe("Render document", function () {
  const tests = [
    'test1',
    'test2'
  ];

  for (let path of tests) {
    it(`from ${path} should be correct`, async () => {

      const docBlob = await fetch(`/base/tests/${path}/document.docx`).then(r => r.blob());
      const resultText = await fetch(`/base/tests/${path}/result.html`).then(r => r.text());

      const div = document.createElement("div");

      document.body.appendChild(div);

      await docx.renderAsync(docBlob, div);
      
      const actual = cleanUpText(div.innerHTML);
      const expected = cleanUpText(resultText);

      expect(actual == expected).toBeTrue();

      if(actual != expected) {
        const diffs = Diff.diffChars(actual, expected);

        for(const diff of diffs) {
          if(diff.added)
            console.log(diff.value);

          if(diff.removed)
            console.error(diff.value);
        }
      }

      div.remove();
    });
  }
});

function cleanUpText(text) {
  return text.replace(/\t+|\s+/ig, ' ');
}