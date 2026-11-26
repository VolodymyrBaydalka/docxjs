describe("universal measures", function () {
  it("reads in and cm lengths as points and a line without lineRule as auto", async () => {
    const zip = new JSZip();
    zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:ind w:left="1in"/></w:pPr></w:p>
<w:p><w:pPr><w:ind w:left="1cm"/></w:pPr></w:p>
<w:p><w:pPr><w:spacing w:line="0.5in" w:lineRule="exact"/></w:pPr></w:p>
<w:p><w:pPr><w:spacing w:line="480"/></w:pPr></w:p></w:body></w:document>`);
    const div = document.createElement("div");
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), div);
    const [inch, cm, exact, noRule] = div.querySelectorAll("article p");
    expect(parseFloat(inch.style.marginInlineStart)).toBeCloseTo(72, 2);
    expect(parseFloat(cm.style.marginInlineStart)).toBeCloseTo(28.35, 2);
    expect(exact.style.lineHeight).toBe("36pt");
    expect(noRule.style.lineHeight).toBe("2");
  });
});
