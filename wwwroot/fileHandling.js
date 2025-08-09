window.downloadFile = (filename, base64Data) => {
    try {
        // Determine MIME type and file type info based on extension
        const fileInfo = getFileTypeInfo(filename);
        
        // Create blob from base64 data for better mobile compatibility
        const byteCharacters = atob(base64Data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: fileInfo.mimeType });
        
        // Use modern download API if available
        if ('showSaveFilePicker' in window) {
            // Modern File System Access API (Chrome/Edge)
            const pickerOptions = {
                suggestedName: filename,
                types: [{
                    description: fileInfo.description,
                    accept: { [fileInfo.mimeType]: [fileInfo.extension] }
                }]
            };
            
            window.showSaveFilePicker(pickerOptions).then(fileHandle => {
                return fileHandle.createWritable();
            }).then(writable => {
                writable.write(blob);
                return writable.close();
            }).catch(err => {
                // Fallback to blob URL method
                fallbackDownload(blob, filename);
            });
        } else {
            // Fallback for other browsers
            fallbackDownload(blob, filename);
        }
    } catch (error) {
        console.error('Download error:', error);
        alert('Download failed. Please try again.');
    }
};

function getFileTypeInfo(filename) {
    const extension = filename.toLowerCase().split('.').pop();
    
    switch (extension) {
        case 'xlsx':
            return {
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                extension: '.xlsx',
                description: 'Excel files'
            };
        case 'zip':
            return {
                mimeType: 'application/zip',
                extension: '.zip',
                description: 'ZIP files'
            };
        case 'json':
            return {
                mimeType: 'application/json',
                extension: '.json',
                description: 'JSON files'
            };
        default:
            return {
                mimeType: 'application/octet-stream',
                extension: '.' + extension,
                description: 'Files'
            };
    }
}

function fallbackDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Clean up the URL object after a delay
    setTimeout(() => URL.revokeObjectURL(url), 100);
}

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