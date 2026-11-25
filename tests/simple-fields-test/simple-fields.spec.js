describe("Simple fields", function () {
    let container;

    beforeEach(() => {
        container = document.createElement("div");
        document.body.appendChild(container);
    });

    afterEach(() => container.remove());

    async function renderBody(body) {
        const zip = new JSZip();
        zip.file("[Content_Types].xml", `
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>`);
        zip.file("_rels/.rels", `
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>`);
        zip.file("word/document.xml", `
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>${body}</w:body>
            </w:document>`);

        await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
    }

    it("renders cached runs in order with their formatting", async () => {
        await renderBody(`
            <w:p>
                <w:r><w:t xml:space="preserve">Before </w:t></w:r>
                <w:fldSimple w:instr=" MERGEFIELD Project.Name ">
                    <w:r><w:rPr><w:b/></w:rPr><w:t>Project</w:t></w:r>
                    <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> Name</w:t></w:r>
                </w:fldSimple>
                <w:r><w:t xml:space="preserve"> After</w:t></w:r>
            </w:p>`);

        const paragraph = container.querySelector("article p");
        const runs = Array.from(paragraph.querySelectorAll("span"));

        expect(paragraph.textContent).toBe("Before Project Name After");
        expect(runs.find(r => r.textContent === "Project")?.style.fontWeight).toBe("bold");
        expect(runs.find(r => r.textContent === " Name")?.style.fontStyle).toBe("italic");
    });

    it("renders cached results inside table cells", async () => {
        await renderBody(`
            <w:tbl><w:tr><w:tc><w:p>
                <w:r><w:t xml:space="preserve">Total: </w:t></w:r>
                <w:fldSimple w:instr=" MERGEFIELD Total ">
                    <w:r><w:t>42</w:t></w:r>
                </w:fldSimple>
                <w:r><w:t xml:space="preserve"> EUR</w:t></w:r>
            </w:p></w:tc></w:tr></w:tbl>`);

        expect(container.querySelector("td").textContent).toBe("Total: 42 EUR");
    });
});
