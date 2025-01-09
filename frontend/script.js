document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const pdfUploadArea = document.getElementById('pdfUploadArea');
    const wordUploadArea = document.getElementById('wordUploadArea');
    const pdfInput = document.getElementById('pdfInput');
    const wordInput = document.getElementById('wordInput');
    const convertToWordBtn = document.getElementById('convertToWord');
    const convertToPdfBtn = document.getElementById('convertToPdf');

    // 添加加载状态指示器
    function showLoading(button) {
        button.disabled = true;
        button.textContent = '转换中...';
    }

    function hideLoading(button, originalText) {
        button.disabled = false;
        button.textContent = originalText;
    }

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

    // 获取进度显示元素
    const pdfProgress = document.getElementById('pdfProgress');
    const wordProgress = document.getElementById('wordProgress');

    // 更新进度显示函数
    function updateProgress(progressElement, message) {
        progressElement.textContent = message;
        progressElement.classList.add('active');
    }

    // 清除进度显示
    function clearProgress(progressElement) {
        progressElement.textContent = '';
        progressElement.classList.remove('active');
    }

    // 修改SSE监听函数
    function listenToProgress(queueId, progressElement) {
        return new Promise((resolve, reject) => {
            let retryCount = 0;
            const maxRetries = 3;
            let eventSource = null;
            
            function connect() {
                if (eventSource) {
                    eventSource.close();
                }

                console.log(`建立SSE连接，队列ID: ${queueId}`);
                eventSource = new EventSource(`http://localhost:5000/progress/${queueId}`);
                let heartbeatTimeout;

                function resetHeartbeatTimeout() {
                    if (heartbeatTimeout) clearTimeout(heartbeatTimeout);
                    heartbeatTimeout = setTimeout(() => {
                        console.log('心跳超时，关闭连接');
                        eventSource.close();
                        if (retryCount < maxRetries) {
                            retryCount++;
                            console.log(`重试连接 ${retryCount}/${maxRetries}`);
                            setTimeout(connect, 1000);
                        } else {
                            reject(new Error('服务器连接超时'));
                        }
                    }, 35000);
                }

                eventSource.onopen = function() {
                    console.log('SSE连接已建立');
                    resetHeartbeatTimeout();
                };

                eventSource.onmessage = function(event) {
                    try {
                        console.log('收到消息:', event.data);
                        const data = JSON.parse(event.data);
                        
                        if (data.error) {
                            throw new Error(data.error);
                        }

                        if (data.status === 'heartbeat') {
                            console.log('收到心跳包');
                            resetHeartbeatTimeout();
                            return;
                        }

                        if (data.status) {
                            updateProgress(progressElement, data.status);
                            if (data.status === 'DONE') {
                                console.log('转换完成');
                                eventSource.close();
                                if (heartbeatTimeout) clearTimeout(heartbeatTimeout);
                                resolve();
                                return;
                            }
                        }
                        
                        resetHeartbeatTimeout();
                    } catch (error) {
                        console.error('处理消息时出错:', error);
                        eventSource.close();
                        if (heartbeatTimeout) clearTimeout(heartbeatTimeout);
                        reject(error);
                    }
                };

                eventSource.onerror = function(event) {
                    console.error('SSE连接错误:', event);
                    eventSource.close();
                    if (heartbeatTimeout) clearTimeout(heartbeatTimeout);
                    if (retryCount < maxRetries) {
                        retryCount++;
                        console.log(`连接错误，重试 ${retryCount}/${maxRetries}`);
                        setTimeout(connect, 1000);
                    } else {
                        reject(new Error('转换过程中断'));
                    }
                };
            }

            connect();
        });
    }

    // 修改转换为Word按钮点击事件
    convertToWordBtn.addEventListener('click', async function() {
        const file = pdfInput.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            showLoading(convertToWordBtn);
            updateProgress(pdfProgress, '正在转换...');

            const response = await fetch('http://localhost:5000/convert/pdf-to-word', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || '转换失败');
            }

            // 下载转换后的文件
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = file.name.replace('.pdf', '.docx');
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            // 重置上传区域
            pdfUploadArea.querySelector('p').textContent = '点击或拖拽PDF文件到这里';
            pdfInput.value = '';
            
            updateProgress(pdfProgress, '转换完成！');
            setTimeout(() => clearProgress(pdfProgress), 3000);
        } catch (error) {
            updateProgress(pdfProgress, `错误: ${error.message}`);
            console.error('转换错误:', error);
        } finally {
            hideLoading(convertToWordBtn, '转换为Word');
        }
    });

    // 修改转换为PDF按钮点击事件
    convertToPdfBtn.addEventListener('click', async function() {
        const file = wordInput.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);
        formData.append('mode', 'accurate');

        try {
            showLoading(convertToPdfBtn);
            updateProgress(wordProgress, '准备转换...');

            const response = await fetch('http://localhost:5000/convert/word-to-pdf', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || '转换失败');
            }

            updateProgress(wordProgress, '转换完成，准备下载...');
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = file.name.replace(/\.docx?$/, '.pdf');
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            // 重置上传区域
            wordUploadArea.querySelector('p').textContent = '点击或拖拽Word文件到这里';
            wordInput.value = '';
            
            setTimeout(() => clearProgress(wordProgress), 3000);
        } catch (error) {
            updateProgress(wordProgress, `错误: ${error.message}`);
            alert(error.message);
        } finally {
            hideLoading(convertToPdfBtn, '转换为PDF');
        }
    });

    // 拖拽上传功能
    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.add('dragover');
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.remove('dragover');
    }

    function handleDrop(e, input) {
        e.preventDefault();
        e.stopPropagation();
        e.currentTarget.classList.remove('dragover');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            input.files = files;
            const event = new Event('change');
            input.dispatchEvent(event);
        }
    }

    // 为上传区域添加拖拽事件
    [pdfUploadArea, wordUploadArea].forEach(area => {
        area.addEventListener('dragover', handleDragOver);
        area.addEventListener('dragleave', handleDragLeave);
    });
    
    pdfUploadArea.addEventListener('drop', (e) => handleDrop(e, pdfInput));
    wordUploadArea.addEventListener('drop', (e) => handleDrop(e, wordInput));
}); 