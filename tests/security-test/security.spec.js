describe("Untrusted document content", function () {
  const relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const containers = [];

  afterEach(() => {
    containers.forEach(container => container.remove());
    containers.length = 0;
    delete window.__docxAltChunkExecuted;
  });

  async function render(body, relationships, parts = {}) {
    const source = await fetch("/base/tests/render-test/text/document.docx").then(response => response.blob());
    const zip = await JSZip.loadAsync(source);

    zip.file("word/document.xml", `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${relationshipType}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
    zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`);
    for (const [path, content] of Object.entries(parts)) {
      zip.file(path, content);
    }

    const container = document.createElement("div");
    document.body.appendChild(container);
    containers.push(container);
    await docx.renderAsync(await zip.generateAsync({ type: "blob" }), container);
    return container;
  }

  it("omits unsafe hyperlink targets while keeping safe links", async () => {
    const targets = [
      ["javascript", "javascript:alert(1)"],
      ["mixed", "JaVaScRiPt:alert(1)"],
      ["newline", "java\nscript:alert(1)"],
      ["leadingTab", "\tjavascript:alert(1)"],
      ["data", "data:text/html,hello"],
      ["vbscript", "vbscript:alert(1)"],
      ["file", "file:///etc/passwd"],
      ["malformed", "http://["],
      ["https", "https://example.com/path"],
      ["mailto", "mailto:test@example.com"],
      ["relative", "../guide"]
    ];
    const relationships = targets.map(([id, target]) =>
      `<Relationship Id="${id}" Type="${relationshipType}/hyperlink" Target="${target.replace(/\n/g, "&#10;").replace(/\t/g, "&#9;")}" TargetMode="External"/>`
    ).join("");
    const body = targets.map(([id]) =>
      `<w:p><w:hyperlink r:id="${id}"><w:r><w:t>${id}</w:t></w:r></w:hyperlink></w:p>`
    ).join("") + '<w:p><w:hyperlink w:anchor="section"><w:r><w:t>anchor</w:t></w:r></w:hyperlink></w:p>';

    const container = await render(body, relationships);
    const links = Array.from(container.querySelectorAll("a"));

    expect(links.length).toBe(targets.length + 1);
    for (let i = 0; i < 8; i++) {
      expect(links[i].hasAttribute("href")).withContext(targets[i][0]).toBe(false);
    }
    expect(links[8].getAttribute("href")).toBe("https://example.com/path");
    expect(links[9].getAttribute("href")).toBe("mailto:test@example.com");
    expect(links[10].getAttribute("href")).toBe("../guide");
    expect(links[11].getAttribute("href")).toBe("#section");
  });

  it("sandboxes altChunk HTML so its scripts cannot reach the host page", async () => {
    const relationships = `<Relationship Id="chunk" Type="${relationshipType}/aFChunk" Target="chunk.html"/>`;
    const container = await render('<w:altChunk r:id="chunk"/>', relationships, {
      "word/chunk.html": "<p>Embedded text</p><script>parent.__docxAltChunkExecuted = true</script>"
    });
    const frame = container.querySelector("iframe");

    expect(frame).not.toBeNull();
    expect(frame.hasAttribute("sandbox")).toBe(true);
    expect(frame.getAttribute("sandbox")).toBe("");
    await new Promise(resolve => frame.addEventListener("load", resolve, { once: true }));
    expect(window.__docxAltChunkExecuted).toBeUndefined();
    expect(frame.srcdoc).toContain("Embedded text");
  });
});
