describe("Numbering instances", function () {
    const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    async function parseNumbering(contents) {
        const zip = new JSZip();
        zip.file("[Content_Types].xml", `
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>`);
        zip.file("_rels/.rels", `
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/>
            </Relationships>`);
        zip.file("word/document.xml", `<w:document xmlns:w="${w}"><w:body><w:p/></w:body></w:document>`);
        zip.file("word/_rels/document.xml.rels", `
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="numbering" Type="${r}/numbering" Target="numbering.xml"/>
            </Relationships>`);
        zip.file("word/numbering.xml", `<w:numbering xmlns:w="${w}">${contents}</w:numbering>`);
        const parsed = await docx.parseAsync(await zip.generateAsync({ type: "uint8array" }));
        return parsed.numberingPart.domNumberings;
    }

    it("preserves instance order, level order and independent instance IDs", async () => {
        const levels = await parseNumbering(`
            <w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num>
            <w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
            <w:num w:numId="9"><w:abstractNumId w:val="1"/></w:num>
            <w:abstractNum w:abstractNumId="2">
                <w:lvl w:ilvl="0"><w:lvlText w:val="%1)"/></w:lvl>
            </w:abstractNum>
            <w:abstractNum w:abstractNumId="1">
                <w:lvl w:ilvl="0"><w:lvlText w:val="%1."/></w:lvl>
                <w:lvl w:ilvl="1"><w:lvlText w:val="%1.%2."/></w:lvl>
            </w:abstractNum>`);

        expect(levels.map(level => [level.id, level.level, level.levelText])).toEqual([
            ["7", 0, "%1."], ["7", 1, "%1.%2."], ["3", 0, "%1)"],
            ["9", 0, "%1."], ["9", 1, "%1.%2."]
        ]);
        expect(levels[0]).not.toBe(levels[3]);
        expect(levels[1]).not.toBe(levels[4]);
    });

    it("omits unused definitions and ignores unresolved abstract references", async () => {
        const definitions = `
            <w:abstractNum w:abstractNumId="1">
                <w:lvl w:ilvl="0"><w:lvlText w:val="%1."/></w:lvl>
            </w:abstractNum>`;
        expect(await parseNumbering(definitions)).toEqual([]);
        expect(await parseNumbering(definitions + `
            <w:num w:numId="5"><w:abstractNumId w:val="999"/></w:num>`)).toEqual([]);
    });
});
