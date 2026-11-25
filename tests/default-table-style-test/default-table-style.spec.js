describe("default table style", function () {
  it("puts its cell margins on the cells", async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    zip.file("word/styles.xml", `<w:styles ${W}><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:left w:w="216" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
    zip.file("word/document.xml", `<w:document ${W}><w:body><w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`);

    const container = document.body.appendChild(document.createElement("div"));
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { className: "dss" });
    // Chrome resolves pt to px: 216 dxa = 10.8pt = 14.4px
    expect(parseFloat(getComputedStyle(container.querySelector("article td")).paddingLeft)).toBeCloseTo(14.4, 2);
    container.remove();
  });
});
