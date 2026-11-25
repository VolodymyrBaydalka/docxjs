describe("Image loading", function () {
    const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let container, dataURLSpy, objectURLSpy;

    beforeEach(() => {
        container = document.createElement("div");
        document.body.appendChild(container);
        dataURLSpy = spyOn(FileReader.prototype, "readAsDataURL").and.callThrough();
        objectURLSpy = spyOn(URL, "createObjectURL").and.callThrough();
    });

    afterEach(() => {
        container.remove();
        for (const call of objectURLSpy.calls.all()) {
            URL.revokeObjectURL(call.returnValue);
        }
    });

    function relationships(body) {
        return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    }

    function relationship(id, type, target) {
        return `<Relationship Id="${id}" Type="${r}/${type}" Target="${target}"/>`;
    }

    async function makeDocument(body = "<w:p/>") {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", `
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="png" ContentType="image/png"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>`);
        zip.file("_rels/.rels", relationships(relationship("main", "officeDocument", "word/document.xml")));
        zip.file("word/document.xml", `
            <w:document xmlns:w="${w}" xmlns:r="${r}"
                xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <w:body>${body}</w:body>
            </w:document>`);
        zip.file("word/_rels/document.xml.rels", relationships(
            relationship("image", "image", "media/shared.png") +
            relationship("alias", "image", "media/../media/shared.png") +
            relationship("header", "header", "header1.xml") +
            relationship("numbering", "numbering", "numbering.xml")
        ));
        zip.file("word/header1.xml", `<w:hdr xmlns:w="${w}"><w:p/></w:hdr>`);
        zip.file("word/_rels/header1.xml.rels", relationships(
            relationship("image", "image", "media/other.png") +
            relationship("shared", "image", "/word/media/shared.png")
        ));
        zip.file("word/numbering.xml", `<w:numbering xmlns:w="${w}"/>`);
        zip.file("word/_rels/numbering.xml.rels", relationships(relationship("image", "image", "media/shared.png")));
        zip.file("word/media/shared.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==", { base64: true });
        zip.file("word/media/other.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGP4DwQACfsD/fteaysAAAAASUVORK5CYII=", { base64: true });
        return zip.generateAsync({ type: "uint8array" });
    }

    function drawing(id) {
        return `<w:p><w:r><w:drawing><wp:inline>
            <wp:extent cx="9525" cy="9525"/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                    <pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm></pic:spPr>
                </pic:pic>
            </a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p>`;
    }

    for (const useBase64URL of [false, true]) {
        it(`converts repeated image placements once (useBase64URL=${useBase64URL})`, async () => {
            const body = Array.from({ length: 8 }, (_, i) => drawing(i % 2 ? "alias" : "image")).join("");
            await docx.renderAsync(await makeDocument(body), container, null, { useBase64URL });

            const images = Array.from(container.querySelectorAll("img"));
            expect(images.length).toBe(8);
            expect(new Set(images.map(image => image.src)).size).toBe(1);
            expect((useBase64URL ? dataURLSpy : objectURLSpy).calls.count()).toBe(1);
        });
    }

    it("shares resolved image paths across parts, but keeps documents and different images separate", async () => {
        const bytes = await makeDocument();
        const options = { useBase64URL: true };
        const first = await docx.parseAsync(bytes, options);
        const second = await docx.parseAsync(bytes, options);
        const header = first.partsMap["word/header1.xml"];
        const urls = await Promise.all([
            first.loadDocumentImage("image"),
            first.loadDocumentImage("alias"),
            first.loadNumberingImage("image"),
            first.loadDocumentImage("shared", header),
            first.loadDocumentImage("image", header),
            second.loadDocumentImage("image")
        ]);

        expect(urls.slice(0, 4)).toEqual(Array(4).fill(urls[0]));
        expect(urls[4]).not.toBe(urls[0]);
        expect(urls[5]).toBe(urls[0]);
        expect(dataURLSpy.calls.count()).toBe(3);
        expect(await first.loadDocumentImage("missing")).toBeNull();
    });

    it("allows retrying an image after conversion fails", async () => {
        const parsed = await docx.parseAsync(await makeDocument(), { useBase64URL: true });
        dataURLSpy.and.throwError("Conversion failed");
        await expectAsync(parsed.loadDocumentImage("image")).toBeRejectedWithError("Conversion failed");

        dataURLSpy.and.callThrough();
        const url = await parsed.loadDocumentImage("image");
        expect(url).toMatch(/^data:image\/png;base64,/);
        expect(await parsed.loadDocumentImage("image")).toBe(url);
        expect(dataURLSpy.calls.count()).toBe(2);
    });
});
