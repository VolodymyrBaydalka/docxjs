describe("Picture bullet styles", function () {
    const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let container;

    beforeEach(() => {
        container = document.createElement("div");
        document.body.appendChild(container);
    });

    afterEach(() => {
        for (const url of new Set(styles().match(/blob:[^)\s"']+/g) || [])) {
            URL.revokeObjectURL(url);
        }
        container.remove();
    });

    function styles() {
        return Array.from(container.querySelectorAll("style"), style => style.textContent).join("\n");
    }

    function relationships(body) {
        return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    }

    function relationship(id, type, target) {
        return `<Relationship Id="${id}" Type="${r}/${type}" Target="${target}"/>`;
    }

    async function renderBullets(bullets, instances, useBase64URL, imageRelationships = true) {
        const zip = new JSZip();
        const pictures = bullets.map(b => `
            <w:numPicBullet w:numPicBulletId="${b.id}"><w:pict>
                <v:shape style="width:${b.width}pt;height:${b.width}pt;">
                    <v:imagedata r:id="${b.src}"/>
                </v:shape>
            </w:pict></w:numPicBullet>`).join("");
        const abstracts = bullets.map(b => `
            <w:abstractNum w:abstractNumId="${b.id}">
                <w:lvl w:ilvl="0"><w:lvlPicBulletId w:val="${b.id}"/></w:lvl>
            </w:abstractNum>`).join("");
        const numbers = instances.map((id, i) => `
            <w:num w:numId="${i + 1}"><w:abstractNumId w:val="${id}"/></w:num>`).join("");
        const body = instances.map((_, i) => `
            <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${i + 1}"/></w:numPr></w:pPr>
                <w:r><w:t>Item ${i + 1}</w:t></w:r>
            </w:p>`).join("");

        zip.file("[Content_Types].xml", `
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Default Extension="png" ContentType="image/png"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>`);
        zip.file("_rels/.rels", relationships(relationship("main", "officeDocument", "word/document.xml")));
        zip.file("word/document.xml", `<w:document xmlns:w="${w}"><w:body>${body}</w:body></w:document>`);
        zip.file("word/_rels/document.xml.rels", relationships(relationship("numbering", "numbering", "numbering.xml")));
        zip.file("word/numbering.xml", `
            <w:numbering xmlns:w="${w}" xmlns:r="${r}" xmlns:v="urn:schemas-microsoft-com:vml">
                ${pictures}${abstracts}${numbers}
            </w:numbering>`);
        if (imageRelationships) {
            zip.file("word/_rels/numbering.xml.rels", relationships(
                relationship("image", "image", "media/shared.png") +
                relationship("alias", "image", "media/../media/shared.png") +
                relationship("other", "image", "media/other.png")
            ));
        }
        zip.file("word/media/shared.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAX+XDSwAAAABJRU5ErkJggg==", { base64: true });
        zip.file("word/media/other.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGP4DwQACfsD/fteaysAAAAASUVORK5CYII=", { base64: true });

        await docx.renderAsync(await zip.generateAsync({ type: "uint8array" }), container, null, { useBase64URL });
    }

    for (const useBase64URL of [false, true]) {
        it(`emits one image URL for repeated numbering instances (useBase64URL=${useBase64URL})`, async () => {
            await renderBullets([{ id: 1, src: "image", width: 1 }], Array(20).fill(1), useBase64URL);

            const paragraphs = Array.from(container.querySelectorAll("article p"));
            const backgrounds = paragraphs.map(p => getComputedStyle(p, "::before").backgroundImage);
            expect(paragraphs.length).toBe(20);
            paragraphs.forEach((p, i) => {
                expect(p.textContent).toBe(`Item ${i + 1}`);
                expect(p.classList.contains(`docx-num-${i + 1}-0`)).toBeTrue();
                expect(getComputedStyle(p).display).toBe("list-item");
            });
            expect(backgrounds[0]).toMatch(/^url\(/);
            expect(new Set(backgrounds).size).toBe(1);
            expect((styles().match(useBase64URL ? /data:image\/png;base64,/g : /blob:/g) || []).length).toBe(1);
        });
    }

    it("shares image aliases while preserving individual bullet sizes and distinct images", async () => {
        await renderBullets([
            { id: 1, src: "image", width: 3 },
            { id: 2, src: "alias", width: 6 },
            { id: 3, src: "other", width: 9 }
        ], [1, 2, 3], true);

        const bullets = Array.from(container.querySelectorAll("article p"), p => getComputedStyle(p, "::before"));
        expect(bullets.length).toBe(3);
        expect(bullets[0].backgroundImage).toMatch(/^url\(/);
        expect(bullets[1].backgroundImage).toBe(bullets[0].backgroundImage);
        expect(bullets[2].backgroundImage).not.toBe(bullets[0].backgroundImage);
        bullets.forEach((bullet, i) => expect(parseFloat(bullet.width)).toBe((i + 1) * 4));
        expect((styles().match(/data:image\/png;base64,/g) || []).length).toBe(2);
    });

    it("keeps rendering list text when picture relationships are missing", async () => {
        await renderBullets([{ id: 1, src: "image", width: 3 }], [1], true, false);

        const paragraph = container.querySelector("article p");
        expect(paragraph.textContent).toBe("Item 1");
        expect(getComputedStyle(paragraph).display).toBe("list-item");
    });
});
