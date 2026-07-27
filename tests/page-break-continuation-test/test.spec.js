describe("page-break numbering continuation", function () {
  const SUPPRESSED = "docx-numbering-suppressed";

  async function render(name) {
    const docBlob = await fetch(`/base/tests/page-break-continuation-test/${name}`).then(r => r.blob());
    const div = document.createElement("div");

    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div, null, {
      ignoreLastRenderedPageBreak: false
    });

    return div;
  }

  function paragraphOf(div, text) {
    return Array.from(div.querySelectorAll("p")).find(p => p.textContent.includes(text));
  }

  it("suppresses numbering on the fragment after a mid-paragraph break", async () => {
    const div = await render("document.docx");

    expect(div.querySelectorAll("section.docx").length).toBe(2);

    const numbered = Array.from(div.querySelectorAll("p.docx-num-1-0"));
    expect(numbered.length).toBe(3);

    const suppressed = Array.from(div.querySelectorAll(`p.${SUPPRESSED}`));
    expect(suppressed.length).toBe(1);
    expect(suppressed[0]).toBe(numbered[1]);
    expect(suppressed[0].textContent).toContain("After page break");

    expect(getComputedStyle(numbered[1], "::before").content).toBe("none");
    expect(getComputedStyle(numbered[0], "::before").content).not.toBe("none");
    expect(getComputedStyle(numbered[2], "::before").content).not.toBe("none");
    expect(getComputedStyle(numbered[2], "::before").counterIncrement).toContain("docx-num-1-0");

    div.remove();
  });

  it("keeps nested counters working around a suppressed fragment", async () => {
    const div = await render("document.docx");

    const nested = div.querySelector("p.docx-num-1-1");
    expect(!!nested).toBe(true);
    expect(getComputedStyle(nested, "::before").counterIncrement).toContain("docx-num-1-1");

    expect(getComputedStyle(div.querySelector("p.docx-num-1-0")).counterSet).toContain("docx-num-1-1");
    expect(getComputedStyle(div.querySelector(`p.${SUPPRESSED}`)).counterSet).toBe("none");

    div.remove();
  });

  it("stops the implicit list-item counter on suppressed fragments", async () => {
    const div = await render("document-native-numbering.docx");

    expect(div.querySelectorAll("section.docx").length).toBe(2);

    const numbered = Array.from(div.querySelectorAll("p.docx-num-1-0"));
    expect(numbered.length).toBe(4);

    const suppressed = div.querySelector(`p.${SUPPRESSED}`);
    expect(suppressed).toBe(numbered[2]);

    expect(getComputedStyle(numbered[2]).counterIncrement).toBe("list-item 0");
    expect(getComputedStyle(numbered[0]).counterIncrement).not.toBe("list-item 0");
    expect(getComputedStyle(numbered[1]).counterIncrement).not.toBe("list-item 0");
    expect(getComputedStyle(numbered[3]).counterIncrement).not.toBe("list-item 0");

    div.remove();
  });

  it("keeps the marker with the text when the break precedes all content", async () => {
    const div = await render("document-break-at-start.docx");

    const pages = div.querySelectorAll("section.docx");
    expect(pages.length).toBe(2);

    const suppressed = Array.from(div.querySelectorAll(`p.${SUPPRESSED}`));
    expect(suppressed.length).toBe(1);
    expect(suppressed[0].textContent).toBe("");
    expect(pages[0].contains(suppressed[0])).toBe(true);

    const secondItem = paragraphOf(div, "Second item");
    expect(pages[1].contains(secondItem)).toBe(true);
    expect(secondItem.classList.contains(SUPPRESSED)).toBe(false);
    expect(getComputedStyle(secondItem, "::before").content).not.toBe("none");
    expect(getComputedStyle(secondItem, "::before").counterIncrement).toContain("docx-num-1-0");

    div.remove();
  });
});
