// pdfHandler.js - 处理PDF相关功能

import { safeLog, updateStatus, handleError, copyTextFallback } from './utils.js';

// 直接在应用内预览PDF，支持文本选择
function renderSimplePdfPreview(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    // 设置预览布局
    filePreviewContent.innerHTML = `
        <div style="height:100%; width:100%; display:flex; flex-direction:column;">
            <div style="display:flex; justify-content:space-between; align-items:center; padding:10px; background-color:#f8f8f8; border-bottom:1px solid #ddd;">
                <div>
                    <span id="current-page">第1页</span> / <span id="total-pages">...</span>
                </div>
                <div>
                    <button id="prev-page-btn" class="ms-Button">上一页</button>
                    <button id="next-page-btn" class="ms-Button">下一页</button>
                    <button id="zoom-in-btn" class="ms-Button">放大</button>
                    <button id="zoom-out-btn" class="ms-Button">缩小</button>
                </div>
            </div>
            <div id="pdf-viewer-container" style="flex:1; overflow:auto; display:flex; justify-content:center; align-items:flex-start; background-color:#888; padding:20px; position:relative;">
                <div class="loading"></div>
                <p style="color:white; margin-left:10px;">正在加载PDF...</p>
            </div>
        </div>
    `;
    
    // 获取PDF内容
    const xhr = new XMLHttpRequest();
    xhr.open('GET', `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`, true);
    xhr.responseType = 'arraybuffer';
    xhr.setRequestHeader('Authorization', `Bearer ${tokenInfo.token}`);
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            // 检查PDF.js是否已加载
            if (typeof pdfjsLib === 'undefined') {
                loadPdfJs().then(() => {
                    renderPdfWithTextLayer(xhr.response);
                }).catch(error => {
                    document.getElementById('pdf-viewer-container').innerHTML = `
                        <div class="error-message">
                            <p>无法加载PDF查看器库: ${error.message}</p>
                        </div>
                    `;
                });
            } else {
                renderPdfWithTextLayer(xhr.response);
            }
        } else {
            document.getElementById('pdf-viewer-container').innerHTML = `
                <div class="error-message">
                    <p>PDF文件加载失败 (${xhr.status})</p>
                </div>
            `;
        }
    };
    
    xhr.onerror = function() {
        document.getElementById('pdf-viewer-container').innerHTML = `
            <div class="error-message">
                <p>网络错误，无法加载PDF</p>
            </div>
        `;
    };
    
    xhr.send();
}

// 动态加载PDF.js
function loadPdfJs() {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js';
        script.onload = function() {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
            resolve();
        };
        script.onerror = function() {
            reject(new Error('无法加载PDF.js库'));
        };
        document.body.appendChild(script);
    });
}

// 使用文本层渲染PDF
function renderPdfWithTextLayer(arrayBuffer) {
    const pdfData = new Uint8Array(arrayBuffer);
    
    // 加载PDF文档
    pdfjsLib.getDocument({data: pdfData}).promise.then(pdf => {
        // 保存PDF引用和当前状态
        window.pdfDocument = pdf;
        window.currentPage = 1;
        window.currentScale = 1.0;
        
        // 更新页数信息
        document.getElementById('total-pages').textContent = `共${pdf.numPages}页`;
        
        // 渲染当前页面
        renderPdfPageWithText(pdf, window.currentPage, window.currentScale);
        
        // 设置翻页和缩放控件
        setupPdfControls(pdf);
    }).catch(error => {
        document.getElementById('pdf-viewer-container').innerHTML = `
            <div class="error-message">
                <p>PDF加载失败: ${error.message}</p>
            </div>
        `;
    });
}

// 渲染带文本层的PDF页面
function renderPdfPageWithText(pdf, pageNumber, scale) {
    const container = document.getElementById('pdf-viewer-container');
    
    // 更新页码显示
    document.getElementById('current-page').textContent = `第${pageNumber}页`;
    
    // 创建页面容器
    const pageContainer = document.createElement('div');
    pageContainer.id = 'pdf-page-container';
    pageContainer.style.position = 'relative';
    pageContainer.style.margin = '0 auto';
    
    // 先清空容器
    container.innerHTML = '';
    container.appendChild(pageContainer);
    
    // 添加加载指示器
    const loadingIndicator = document.createElement('div');
    loadingIndicator.className = 'loading';
    loadingIndicator.style.position = 'absolute';
    loadingIndicator.style.top = '50%';
    loadingIndicator.style.left = '50%';
    loadingIndicator.style.transform = 'translate(-50%, -50%)';
    container.appendChild(loadingIndicator);
    
    // 获取页面
    pdf.getPage(pageNumber).then(page => {
        // 计算页面尺寸
        const viewport = page.getViewport({scale: scale});
        
        // 创建canvas
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.display = 'block';
        canvas.style.border = '1px solid #ddd';
        canvas.style.margin = '0 auto';
        canvas.style.backgroundColor = 'white';
        
        // 添加到容器
        pageContainer.style.width = `${viewport.width}px`;
        pageContainer.style.height = `${viewport.height}px`;
        pageContainer.appendChild(canvas);
        
        // 渲染页面到canvas
        const context = canvas.getContext('2d');
        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };
        
        // 创建文本层div
        const textLayerDiv = document.createElement('div');
        textLayerDiv.className = 'textLayer';
        textLayerDiv.style.position = 'absolute';
        textLayerDiv.style.left = '0';
        textLayerDiv.style.top = '0';
        textLayerDiv.style.right = '0';
        textLayerDiv.style.bottom = '0';
        textLayerDiv.style.overflow = 'hidden';
        textLayerDiv.style.opacity = '0.2';
        textLayerDiv.style.lineHeight = '1.0';
        textLayerDiv.style.userSelect = 'text';
        textLayerDiv.style.cursor = 'text';
        pageContainer.appendChild(textLayerDiv);
        
        // 添加CSS样式
        let style = document.getElementById('pdf-text-layer-style');
        if (!style) {
            style = document.createElement('style');
            style.id = 'pdf-text-layer-style';
            style.textContent = `
                .textLayer {
                    position: absolute;
                    left: 0;
                    top: 0;
                    right: 0;
                    bottom: 0;
                    overflow: hidden;
                    opacity: 0.2;
                    line-height: 1.0;
                }
                .textLayer > span {
                    color: transparent;
                    position: absolute;
                    white-space: pre;
                    cursor: text;
                    transform-origin: 0% 0%;
                }
                .textLayer .highlight {
                    margin: -1px;
                    padding: 1px;
                    background-color: #b400aa;
                    border-radius: 4px;
                }
                .textLayer .highlight.active {
                    background-color: #ff9632;
                }
                .textLayer ::selection {
                    background: #00f;
                    color: transparent;
                }
            `;
            document.head.appendChild(style);
        }
        
        // 渲染页面
        const renderTask = page.render(renderContext);
        renderTask.promise.then(() => {
            // 渲染文本层
            return page.getTextContent().then(textContent => {
                // 创建文本层
                const renderTextLayer = function(textContent, viewport) {
                    textLayerDiv.innerHTML = '';
                    textLayerDiv.style.opacity = '0.2';
                    
                    textContent.items.forEach(function(item) {
                        const tx = pdfjsLib.Util.transform(
                            pdfjsLib.Util.transform(viewport.transform, item.transform),
                            [1, 0, 0, -1, 0, 0]
                        );
                        
                        const style = `
                            left: ${tx[4]}px;
                            top: ${tx[5]}px;
                            font-size: ${tx[0]}px;
                            font-family: sans-serif;
                            transform: scaleX(${tx[0] / tx[0]});
                        `;
                        
                        const textSpan = document.createElement('span');
                        textSpan.textContent = item.str;
                        textSpan.style.cssText = style;
                        textLayerDiv.appendChild(textSpan);
                    });
                };
                
                renderTextLayer(textContent, viewport);
                
                // 移除加载指示器
                if (loadingIndicator && loadingIndicator.parentNode) {
                    loadingIndicator.parentNode.removeChild(loadingIndicator);
                }
            });
        }).catch(error => {
            console.error('渲染PDF页面失败:', error);
            
            // 移除加载指示器，显示错误
            if (loadingIndicator && loadingIndicator.parentNode) {
                loadingIndicator.parentNode.removeChild(loadingIndicator);
            }
            
            container.innerHTML = `
                <div class="error-message">
                    <p>渲染PDF页面失败: ${error.message}</p>
                </div>
            `;
        });
    }).catch(error => {
        console.error('获取PDF页面失败:', error);
        
        // 移除加载指示器，显示错误
        if (loadingIndicator && loadingIndicator.parentNode) {
            loadingIndicator.parentNode.removeChild(loadingIndicator);
        }
        
        container.innerHTML = `
            <div class="error-message">
                <p>获取PDF页面失败: ${error.message}</p>
            </div>
        `;
    });
}

// 设置PDF控件
function setupPdfControls(pdf) {
    document.getElementById('prev-page-btn').onclick = function() {
        if (window.currentPage > 1) {
            window.currentPage--;
            renderPdfPageWithText(pdf, window.currentPage, window.currentScale);
        }
    };
    
    document.getElementById('next-page-btn').onclick = function() {
        if (window.currentPage < pdf.numPages) {
            window.currentPage++;
            renderPdfPageWithText(pdf, window.currentPage, window.currentScale);
        }
    };
    
    document.getElementById('zoom-in-btn').onclick = function() {
        window.currentScale = Math.min(window.currentScale * 1.25, 3.0);
        renderPdfPageWithText(pdf, window.currentPage, window.currentScale);
    };
    
    document.getElementById('zoom-out-btn').onclick = function() {
        window.currentScale = Math.max(window.currentScale * 0.8, 0.5);
        renderPdfPageWithText(pdf, window.currentPage, window.currentScale);
    };
}

// 复制当前PDF页面文本
function copyCurrentPageText(pdf, pageNumber) {
    // 显示加载指示
    const copyBtn = document.getElementById('copy-page-text-btn');
    const originalText = copyBtn.textContent;
    copyBtn.textContent = '正在提取文本...';
    copyBtn.disabled = true;
    
    // 获取页面文本内容
    pdf.getPage(pageNumber).then(function(page) {
        return page.getTextContent();
    }).then(function(textContent) {
        // 处理和组织文本内容
        let lastY = null;
        let text = '';
        
        // 按行组织文本
        textContent.items.forEach(item => {
            // 检查Y坐标变化来确定是否是新行
            if (lastY !== null && Math.abs(lastY - item.transform[5]) > 5) {
                text += '\n'; // 不同行添加换行
            } else if (lastY !== null && item.transform[5] === lastY && text.charAt(text.length - 1) !== ' ') {
                text += ' '; // 同一行确保有空格
            }
            text += item.str;
            lastY = item.transform[5];
        });
        
        // 复制到剪贴板
        navigator.clipboard.writeText(text).then(function() {
            // 复制成功
            copyBtn.textContent = '✓ 复制成功';
            setTimeout(() => {
                copyBtn.textContent = originalText;
                copyBtn.disabled = false;
            }, 2000);
        }).catch(function(err) {
            // 剪贴板API失败，尝试备用方法
            copyTextFallback(text);
            copyBtn.textContent = originalText;
            copyBtn.disabled = false;
        });
    }).catch(function(error) {
        // 获取文本内容失败
        copyBtn.textContent = '✗ 复制失败';
        setTimeout(() => {
            copyBtn.textContent = originalText;
            copyBtn.disabled = false;
        }, 2000);
        
        console.error('获取PDF文本失败:', error);
        alert('无法提取PDF文本: ' + error.message);
    });
}

// 提取PDF文本
function extractPdfText(fileId) {
    const textContainer = document.getElementById('pdf-text-content');
    textContainer.style.display = 'block';
    textContainer.innerHTML = '<div class="loading"></div><p>正在提取文本...</p>';
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        file_id: fileId
    };
    
    // 请求后端API
    fetch(`${API_BASE_URL}/api/extract-text`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`服务器错误: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        if (data.text) {
            textContainer.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <span style="font-weight:bold;">提取的文本内容:</span>
                    <button id="copy-pdf-text" class="ms-Button ms-Button--primary">复制文本</button>
                </div>
                <div style="max-height:300px; overflow:auto; background-color:#f9f9f9; padding:15px; border-radius:4px; border:1px solid #eee;">
                    <pre style="white-space:pre-wrap; font-family:Arial, sans-serif; margin:0;">${data.text}</pre>
                </div>
            `;
            
            // 添加复制按钮功能
            document.getElementById('copy-pdf-text').addEventListener('click', function() {
                navigator.clipboard.writeText(data.text)
                    .then(() => alert('文本已复制到剪贴板'))
                    .catch(err => {
                        console.error('复制失败:', err);
                        copyTextFallback(data.text);
                    });
            });
        } else {
            throw new Error('未能提取文本内容');
        }
    })
    .catch(error => {
        textContainer.innerHTML = `
            <div class="error-message">
                提取文本失败: ${error.message}
            </div>
        `;
    });
}

export { renderSimplePdfPreview, loadPdfJs, renderPdfWithTextLayer, renderPdfPageWithText, setupPdfControls, copyCurrentPageText, extractPdfText };