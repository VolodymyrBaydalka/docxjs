[![npm version](https://badge.fury.io/js/docx-preview.svg)](https://www.npmjs.com/package/docx-preview)

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
