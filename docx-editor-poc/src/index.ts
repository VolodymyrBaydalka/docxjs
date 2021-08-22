import { DocxEditorCore } from './docx-editor-core';

declare const saveAs: any; 

const root = document.querySelector("#root") as HTMLElement;
const fileInput = document.querySelector("#fileInput") as HTMLInputElement;
const loadBtn = document.querySelector("#loadBtn") as HTMLButtonElement;
const saveBtn = document.querySelector("#saveBtn") as HTMLButtonElement;

const editor = new DocxEditorCore();
editor.init(root, root);

loadBtn.addEventListener('click', () => {
    if (fileInput.files.length > 0) {
        editor.open(fileInput.files[0]);
    }
});

saveBtn.addEventListener('click', () => {
    editor.save().then(saveAs)
});
