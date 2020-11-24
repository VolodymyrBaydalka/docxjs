describe("Render document", function () {
  let tests = [
    'test1'
  ];

  for (let path of tests) {
    it(`from ${path} should be correct`, async () => {

      let docBlob = await fetch(`/base/tests/${path}/document.docx`).then(r => r.blob());
      let resultText = await fetch(`/base/tests/${path}/result.html`).then(r => r.text());

      let div = document.createElement("div");

      document.body.appendChild(div);

      await docx.renderAsync(docBlob, div);
      
      let actual = cleanUpText(div.innerHTML);
      let expected = cleanUpText(resultText);

      expect(actual == expected).toBeTrue();

      if(actual != expected) {
        let diffs = Diff.diffChars(actual, expected);

        for(let diff of diffs) {
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