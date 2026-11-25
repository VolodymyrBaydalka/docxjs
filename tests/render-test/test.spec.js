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
    'header-footer-even-odd-off',
    'footnote',
    'equation',
    'page-break-before'
  ];

  for (let path of tests) {
    it(`from ${path} should be correct`, async () => {

      const docBlob = await fetch(`/base/tests/render-test/${path}/document.docx`).then(r => r.blob());
      const resultText = await fetch(`/base/tests/render-test/${path}/result.html`).then(r => r.text());

      const div = document.createElement("div");

      document.body.appendChild(div);

      await docx.renderAsync(docBlob, div, null, {
        renderChanges: true
      });
      
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

  it("renders distinct outer and inside table borders", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`);
    zip.file("word/document.xml", `<?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:tbl>
          <w:tblPr><w:tblBorders>
            <w:top w:val="single" w:sz="8" w:color="111111"/>
            <w:left w:val="single" w:sz="8" w:color="111111"/>
            <w:bottom w:val="single" w:sz="8" w:color="111111"/>
            <w:right w:val="single" w:sz="8" w:color="111111"/>
            <w:insideH w:val="dashed" w:sz="16" w:color="222222"/>
            <w:insideV w:val="dotted" w:sz="24" w:color="333333"/>
          </w:tblBorders></w:tblPr>
          <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
          <w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
        </w:tbl></w:body>
      </w:document>`);

    const docBlob = await zip.generateAsync({ type: "blob" });
    const div = document.createElement("div");
    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div);

    const rows = div.querySelectorAll("tr");
    const topLeft = rows[0].cells[0].style;
    const bottomRight = rows[1].cells[1].style;

    expect(topLeft.borderTop).toBe("1pt solid rgb(17, 17, 17)");
    expect(topLeft.borderLeft).toBe("1pt solid rgb(17, 17, 17)");
    expect(topLeft.borderRight).toBe("3pt dotted rgb(51, 51, 51)");
    expect(topLeft.borderBottom).toBe("2pt dashed rgb(34, 34, 34)");
    expect(bottomRight.borderTop).toBe("2pt dashed rgb(34, 34, 34)");
    expect(bottomRight.borderLeft).toBe("3pt dotted rgb(51, 51, 51)");
    expect(bottomRight.borderRight).toBe("1pt solid rgb(17, 17, 17)");
    expect(bottomRight.borderBottom).toBe("1pt solid rgb(17, 17, 17)");

    div.remove();
  });
});

function formatHTML(text) {
  return text.replace(/\t+|\s+/ig, ' ').replace(/></ig, '>\n<');
}
