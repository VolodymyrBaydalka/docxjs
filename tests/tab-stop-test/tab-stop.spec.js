describe("tab stops", function () {
    it("measures a declared stop from the paragraph text area", async () => {
        const docBlob = await fetch(`/base/tests/tab-stop-test/document.docx`).then(r => r.blob());

        const div = document.createElement("div");
        document.body.appendChild(div);

        await docx.renderAsync(docBlob, div, null, { experimental: true });

        // updateTabStop runs on a timer after render
        await new Promise(r => setTimeout(r, 700));

        // the paragraph carries ind left=1440 hanging=720 and a left tab stop at 1440,
        // so the run after the tab must start exactly on the paragraph's own left indent
        // (= its border box). With the origin measured wrong the text lands a full
        // margin (~96px) to the right.
        const p = div.querySelector("section.docx article > p");
        const spans = Array.from(p.querySelectorAll("span")).filter(s => s.textContent.trim());
        const after = spans[spans.length - 1];

        expect(after.textContent).toBe("Heading text after the tab");

        const delta = after.getBoundingClientRect().left - p.getBoundingClientRect().left;
        expect(Math.abs(delta)).toBeLessThan(6);

        div.remove();
    })
})
