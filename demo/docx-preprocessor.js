async function preprocessDocx(blob) {
    let zip = await  JSZip.loadAsync(blob);
    const files = zip.file(/[.](tiff|wmf)?$/);

    if (files.length == 0)
        return blob;

    for (let f of files) {
        const buffer = await f.async("uint8array");

        if (f.name.endsWith(".tiff")) {
            const tiff = new Tiff({ buffer });
            const blob = await new Promise(res => tiff.toCanvas().toBlob(blob => res(blob), "image/png"));
            zip.file(f.name, blob);
        }
        else if (f.name.endsWith(".wmf")) {
            var renderer = new WMFJS.Renderer(buffer);
            var width = 1000;
            var height = 800;
            var res = renderer.render({
                width: width + "px",
                height: height + "px",
                xExt: width,
                yExt: height,
                mapMode: 8 // preserve aspect ratio checkbox
            });
            var svg = res.firstChild;
            svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
            svg.removeAttribute("width");
            svg.removeAttribute("height");
            zip.file(f.name, svg.outerHTML);
        }
    }

    const contentType = await zip.file("[Content_Types].xml");

    if (contentType) {
        const text = await contentType.async("text");
        zip.file("[Content_Types].xml", text.replace(/image\/x-wmf/g, "image/svg+xml"));
    }

    return await zip.generateAsync({ type: "blob" });
}