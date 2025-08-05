window.downloadFile = (filename, base64Data) => {
    const link = document.createElement('a');
    link.href = 'data:application/zip;base64,' + base64Data;
    link.download = filename;
    link.click();
    link.remove();
};

// Enhanced drag and drop support
window.initializeDragDrop = (dotNetReference) => {
    const dropZone = document.querySelector('[data-dropzone]');
    if (!dropZone) return;

    let dragCounter = 0;

    dropZone.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dragCounter++;
        if (dragCounter === 1) {
            dotNetReference.invokeMethodAsync('HandleDragEnter');
        }
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
            dotNetReference.invokeMethodAsync('HandleDragLeave');
        }
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
    });

    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dragCounter = 0;
        
        const files = Array.from(e.dataTransfer.files);
        for (const file of files) {
            const arrayBuffer = await file.arrayBuffer();
            const bytes = new Uint8Array(arrayBuffer);
            await dotNetReference.invokeMethodAsync('HandleDroppedFile', file.name, file.size, bytes);
        }
    });
};