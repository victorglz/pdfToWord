document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const pdfUploadArea = document.getElementById('pdfUploadArea');
    const wordUploadArea = document.getElementById('wordUploadArea');
    const pdfInput = document.getElementById('pdfInput');
    const wordInput = document.getElementById('wordInput');
    const convertToWordBtn = document.getElementById('convertToWord');
    const convertToPdfBtn = document.getElementById('convertToPdf');

    // PDF上传区域点击事件
    pdfUploadArea.addEventListener('click', () => pdfInput.click());
    
    // Word上传区域点击事件
    wordUploadArea.addEventListener('click', () => wordInput.click());

    // PDF文件选择事件
    pdfInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            const file = e.target.files[0];
            if (file.type === 'application/pdf') {
                convertToWordBtn.disabled = false;
                pdfUploadArea.querySelector('p').textContent = `已选择: ${file.name}`;
            } else {
                alert('请选择PDF文件！');
            }
        }
    });

    // Word文件选择事件
    wordInput.addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            const file = e.target.files[0];
            if (file.name.endsWith('.doc') || file.name.endsWith('.docx')) {
                convertToPdfBtn.disabled = false;
                wordUploadArea.querySelector('p').textContent = `已选择: ${file.name}`;
            } else {
                alert('请选择Word文件！');
            }
        }
    });

    // 转换为Word按钮点击事件
    convertToWordBtn.addEventListener('click', function() {
        alert('注意：由于浏览器安全限制，实际的文件转换需要后端服务支持。这里只是界面演示。');
    });

    // 转换为PDF按钮点击事件
    convertToPdfBtn.addEventListener('click', function() {
        alert('注意：由于浏览器安全限制，实际的文件转换需要后端服务支持。这里只是界面演示。');
    });

    // 拖拽上传功能
    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function handleDrop(e, input) {
        e.preventDefault();
        e.stopPropagation();
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            input.files = files;
            const event = new Event('change');
            input.dispatchEvent(event);
        }
    }

    // 为上传区域添加拖拽事件
    pdfUploadArea.addEventListener('dragover', handleDragOver);
    pdfUploadArea.addEventListener('drop', (e) => handleDrop(e, pdfInput));
    
    wordUploadArea.addEventListener('dragover', handleDragOver);
    wordUploadArea.addEventListener('drop', (e) => handleDrop(e, wordInput));
}); 