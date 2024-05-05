async function preprocessTiff(blob) {
    let zip = await  JSZip.loadAsync(blob);
    const tiffs = zip.file(/[.]tiff?$/);

    if (tiffs.length == 0)
        return blob;

    for (let f of tiffs) {
        const buffer = await f.async("uint8array");
        const tiff = new Tiff({ buffer });
        const blob = await new Promise(res => tiff.toCanvas().toBlob(blob => res(blob), "image/png"));
        zip.file(f.name, blob);
    }

    return await zip.generateAsync({ type: "blob" });
}