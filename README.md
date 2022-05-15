[![npm version](https://badge.fury.io/js/docx-preview.svg)](https://www.npmjs.com/package/docx-preview)
[![Support Ukraine](https://img.shields.io/badge/Support-Ukraine-blue?style=flat&logo=adguard)](https://war.ukraine.ua/)

# docxjs
Docx rendering library

Demo - http://zVolodymyr.github.io/docxjs

Usage
-----
```html
<!--optional polyfill for promise-->
<script src="https://unpkg.com/promise-polyfill/dist/polyfill.min.js"></script>
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="docx-preview.min.js"></script>
<script>
    var docData = <document Blob>;

    docx.renderAsync(docData, document.getElementById("container"))
        .then(x => console.log("docx: finished"));
</script>
<body>
    ...
    <div id="container"></div>
    ...
</body>
```
API
---
```ts
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        className: string = "docx", //class name/prefix for default and document style classes
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        ignoreWidth: boolean = false, //disables rendering width of page
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreFonts: boolean = false, //disables fonts rendering
        breakPages: boolean = true, //enables page breaking on page breaks
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        experimental: boolean = false, //enables experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        useMathMLPolyfill: boolean = false, //includes MathML polyfills for chrome, edge, etc.
        debug: boolean = false, //enables additional logging
    }
    ): Promise<any>
```
Status and stability
------
So far I can't come up with final approach of parsing documents and final structure of API. Main development is moved to **next** branch. Only **renderAsync** function is stable and definition shouldn't be changed in future. Inner implementation of parsing and rendering may be changed at any point of time.
