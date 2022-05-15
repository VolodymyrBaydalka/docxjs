function renderThumbnails(docxContainer, thumbnailsContainer) {
    const sections = docxContainer.querySelectorAll('.docx-wrapper>section');

    thumbnailsContainer.innerHTML = "";
    
    for (let i = 0; i < sections.length; i ++) {
        const id = `docx-page-${i + 1}`;
        const thumbnail = document.createElement('a');

        thumbnail.className = 'docx-thumbnail-item';
        thumbnail.href = `#${id}`;
        thumbnail.innerText = `${i + 1}`;
        thumbnailsContainer.appendChild(thumbnail);

        sections[i].setAttribute("id", id);
    }
}