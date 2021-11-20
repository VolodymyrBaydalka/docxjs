describe("extended-props", function () {
    it("loads extended props", async () => {
        let docBlob = await fetch(`/base/tests/extended-props-test/document.docx`).then(r => r.blob());

        let div = document.createElement("div");

        document.body.appendChild(div);
        
        let docParsed = await docx.renderAsync(docBlob, div);

        expect(!!docParsed.extendedPropsPart == true)
        expect(docParsed.extendedPropsPart.appVersion == "16.0000");
        expect(docParsed.extendedPropsPart.application == "Microsoft Office Word");
        expect(docParsed.extendedPropsPart.characters == 393);
        expect(docParsed.extendedPropsPart.company == "");
        expect(docParsed.extendedPropsPart.lines == 3);
        expect(docParsed.extendedPropsPart.pages == 3);
        expect(docParsed.extendedPropsPart.paragraphs == 1);
        expect(docParsed.extendedPropsPart.template == "Normal.dotm");
        expect(docParsed.extendedPropsPart.words == 68);
    })
})