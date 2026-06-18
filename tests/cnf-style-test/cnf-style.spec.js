describe("cnf-style", function () {
    // Builds a minimal docx whose table cell carries a w:cnfStyle without a w:val
    // attribute. Word writes cnfStyle this way for some documents, and before the
    // fix classNameOfCnfStyle dereferenced the missing value and threw
    // "TypeError: Cannot read properties of null (reading '0')".
    function buildDocxBlob() {
        const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

        const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

        const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
<w:tr>
<w:tc>
<w:tcPr><w:cnfStyle w:firstRow="1"/></w:tcPr>
<w:p><w:r><w:t>cell</w:t></w:r></w:p>
</w:tc>
</w:tr>
</w:tbl>
</w:body>
</w:document>`;

        const zip = new JSZip();
        zip.file("[Content_Types].xml", contentTypes);
        zip.folder("_rels").file(".rels", rels);
        zip.folder("word").file("document.xml", document);
        return zip.generateAsync({ type: "blob" });
    }

    it("renders table cells with a cnfStyle that has no val attribute", async () => {
        const docBlob = await buildDocxBlob();

        const div = document.createElement("div");
        document.body.appendChild(div);

        await docx.renderAsync(docBlob, div);

        expect(div.textContent).toContain("cell");

        div.remove();
    });
});
