describe("Tracked change metadata", function () {
  it("exposes author/date/id on rendered <ins> and <del>", async () => {
    const docBlob = await fetch('/base/tests/render-test/revision/document.docx').then(r => r.blob());

    const div = document.createElement("div");
    document.body.appendChild(div);

    await docx.renderAsync(docBlob, div, null, { renderChanges: true });

    const ins = div.querySelector("ins");
    const del = div.querySelector("del");

    expect(ins).not.toBeNull();
    expect(del).not.toBeNull();

    expect(ins.getAttribute("data-change-author")).toBe("Невідомий автор");
    expect(ins.getAttribute("data-change-date")).toBe("2022-09-10T20:45:01Z");
    expect(ins.getAttribute("data-change-id")).toBe("0");

    expect(del.getAttribute("data-change-author")).toBe("Невідомий автор");
    expect(del.getAttribute("data-change-date")).toBe("2022-09-10T20:47:04Z");
    expect(del.getAttribute("data-change-id")).toBe("1");

    div.remove();
  });
});
