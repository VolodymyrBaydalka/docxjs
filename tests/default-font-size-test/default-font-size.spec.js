describe("default font size", function () {
  it("draws text at 10pt when no default sets a size", async () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    // a Normal style and document defaults that set a font but no size
    zip.file("word/styles.xml", `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
    zip.file("word/document.xml", `<w:document ${W}><w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p></w:body></w:document>`);

    const container = document.body.appendChild(document.createElement("div"));
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container, null, { className: "dfs" });
    // Chrome resolves pt to px: 10pt = 13.333px
    const p = container.querySelector("article p");
    expect(parseFloat(getComputedStyle(p.querySelector("span")).fontSize)).toBeCloseTo(13.333, 2);
    expect(parseFloat(getComputedStyle(p).fontSize)).toBeCloseTo(13.333, 2);
    container.remove();
  });
});
