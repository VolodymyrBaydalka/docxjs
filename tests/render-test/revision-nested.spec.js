describe("Nested tracked changes", function () {
  it("renders an <ins> nested inside a <del> (del ⊃ ins) with metadata on both", async () => {
    const docBlob = await fetch('/base/tests/render-test/revision-nested/document.docx').then(r => r.blob());

    const div = document.createElement("div");
    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div, null, { renderChanges: true });

    // The library must preserve the nesting structure, not flatten it.
    const nested = div.querySelector("del > ins");
    expect(nested).not.toBeNull();
    expect(nested.textContent).toBe("nested-insert");

    // Metadata is carried on each level independently.
    const del = div.querySelector("del");
    expect(del.getAttribute("data-change-author")).toBe("Outer Deleter");
    expect(del.getAttribute("data-change-id")).toBe("10");

    expect(nested.getAttribute("data-change-author")).toBe("Inner Inserter");
    expect(nested.getAttribute("data-change-id")).toBe("11");

    div.remove();
  });
});
