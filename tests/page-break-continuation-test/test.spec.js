describe("page-break list continuation", function () {
  it("preserves numbering across lastRenderedPageBreak", async () => {
    const docBlob = await fetch("/base/tests/page-break-continuation-test/document.docx").then(r => r.blob());
    const div = document.createElement("div");
    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div, null, {
      ignoreLastRenderedPageBreak: false
    });

    const pages = div.querySelectorAll("section.docx");
    expect(pages.length).toBe(2);

    const numbered = Array.from(div.querySelectorAll("p.docx-num-1-0"));
    expect(numbered.length).toBe(3);

    const continuations = Array.from(div.querySelectorAll("p.docx-page-break-continuation"));
    expect(continuations.length).toBe(1);
    expect(continuations[0]).toBe(numbered[1]);
    expect(numbered[0].classList.contains("docx-page-break-continuation")).toBe(false);
    expect(numbered[2].classList.contains("docx-page-break-continuation")).toBe(false);

    const continuationStyle = getComputedStyle(continuations[0]);
    expect(continuationStyle.counterSet).toBe("none");
    expect(continuationStyle.listStyleType).toBe("none");

    const continuationBefore = getComputedStyle(continuations[0], "::before");
    expect(continuationBefore.content).toBe("none");
    expect(continuationBefore.counterIncrement).toBe("none");

    const firstBefore = getComputedStyle(numbered[0], "::before");
    const nextBefore = getComputedStyle(numbered[2], "::before");
    expect(firstBefore.content).not.toBe("none");
    expect(nextBefore.content).not.toBe("none");
    expect(nextBefore.counterIncrement).toContain("docx-num-1-0");

    const firstLevel0 = numbered[0];
    expect(getComputedStyle(firstLevel0).counterSet).toContain("docx-num-1-1");
    expect(continuationStyle.counterSet).toBe("none");

    const nested = div.querySelector("p.docx-num-1-1");
    expect(!!nested).toBe(true);
    expect(getComputedStyle(nested, "::before").counterIncrement).toContain("docx-num-1-1");

    div.remove();
  });
});
