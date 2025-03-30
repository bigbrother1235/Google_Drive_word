// documentHandler.js - 处理文档相关功能

import { safeLog, updateStatus, handleError, sanitizeHtml, copyTextFallback } from './utils.js';
import { downloadFile } from './filePreview.js';

// 提取文档文本
function extractDocumentText(fileId, mimeType) {
    // 在预览区域内创建或获取文本容器
    let textContainer = document.getElementById('word-text-container');
    if (!textContainer) {
        // 如果容器不存在，创建一个新的
        const filePreviewContent = document.getElementById('file-preview-content');
        filePreviewContent.innerHTML = `
            <div style="height:100%; width:100%; padding:20px; overflow:auto;">
                <div id="word-text-container" style="background:white; padding:20px; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.1);">
                    <div class="loading"></div>
                    <p>正在提取文本...</p>
                </div>
            </div>
        `;
        textContainer = document.getElementById('word-text-container');
    } else {
        // 如果容器存在，显示加载状态
        textContainer.innerHTML = '<div class="loading"></div><p>正在提取文本...</p>';
        textContainer.style.display = 'block';
    }
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        file_id: fileId,
        mime_type: mimeType
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
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
                    <h3 style="margin:0;">文档文本内容</h3>
                    <button id="copy-extracted-text" class="ms-Button ms-Button--primary">复制文本</button>
                </div>
                <div style="max-height:500px; overflow:auto; background-color:#f9f9f9; padding:20px; border-radius:4px; border:1px solid #eee;">
                    <pre style="white-space:pre-wrap; font-family:Arial, sans-serif; margin:0; line-height:1.5;">${data.text}</pre>
                </div>
            `;
            
            // 添加复制按钮功能
            document.getElementById('copy-extracted-text').addEventListener('click', function() {
                navigator.clipboard.writeText(data.text)
                    .then(() => {
                        this.textContent = '✓ 已复制';
                        setTimeout(() => {
                            this.textContent = '复制文本';
                        }, 2000);
                    })
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
                提取文本失败: ${error.message}<br>
                <button id="try-download-instead" class="ms-Button ms-Button--primary" style="margin-top:12px;">
                    下载文档查看
                </button>
            </div>
        `;
        
        // 添加下载按钮功能
        document.getElementById('try-download-instead').addEventListener('click', function() {
            downloadFile(fileId, '文档.docx');
        });
    });
}

// 简化的Google文档预览
function renderSimpleGoogleDocPreview(file, content) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    if (content) {
        // 安全处理HTML内容
        const sanitizedContent = sanitizeHtml(content);
        
        filePreviewContent.innerHTML = `
            <div style="height:100%; width:100%; overflow:auto; background-color:#f5f5f5; padding:20px;">
                <div style="background-color:white; padding:40px; box-shadow:0 2px 10px rgba(0,0,0,0.1); max-width:800px; margin:0 auto; border-radius:8px;">
                    <div style="display:flex; justify-content:flex-end; margin-bottom:16px;">
                        <button id="copy-doc-text" class="ms-Button ms-Button--primary">复制文本</button>
                    </div>
                    <div class="doc-content" style="font-family:Arial, sans-serif; line-height:1.5;">
                        ${sanitizedContent}
                    </div>
                </div>
            </div>
        `;
        
        // 添加复制按钮功能
        document.getElementById('copy-doc-text').addEventListener('click', function() {
            const docContent = document.querySelector('.doc-content');
            const textContent = docContent.innerText || docContent.textContent;
            
            navigator.clipboard.writeText(textContent)
                .then(() => alert('文本已复制到剪贴板'))
                .catch(err => {
                    console.error('复制失败:', err);
                    copyTextFallback(textContent);
                });
        });
    } else {
        // 如果没有内容，显示备选选项
        filePreviewContent.innerHTML = `
            <div style="text-align:center; padding:30px; background-color:white; border-radius:8px; height:100%; display:flex; flex-direction:column; justify-content:center; align-items:center;">
                <div style="font-size:64px; margin-bottom:20px;">📄</div>
                <h3 style="margin-bottom:20px;">Google 文档</h3>
                <p style="color:#5f6368; margin-bottom:30px;">无法加载文档内容。您可以：</p>
                <div style="display:flex; gap:16px;">
                    <button id="doc-drive-view" class="ms-Button ms-Button--primary">
                        在Google Drive中查看
                    </button>
                    <button id="extract-doc-text" class="ms-Button">
                        提取文档文本
                    </button>
                </div>
                <div id="doc-text-container" style="margin-top:20px; text-align:left; width:100%; display:none;">
                    <div class="loading"></div>
                    <p>正在提取文本...</p>
                </div>
            </div>
        `;
        
        // 在Drive中查看按钮
        document.getElementById('doc-drive-view').addEventListener('click', function() {
            window.open(file.webViewLink, '_blank');
        });
        
        // 提取文本按钮
        document.getElementById('extract-doc-text').addEventListener('click', function() {
            extractDocumentText(file.id, file.mimeType);
        });
    }
}

// 简化的Word文档预览
function renderSimpleWordPreview(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    // 获取文件详细信息以显示在界面上
    const fileName = file.name || '未知文件';
    const fileType = file.mimeType || 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const fileIcon = '📝';
    
    // 直接显示文档信息和选项，不尝试预览内容
    filePreviewContent.innerHTML = `
        <div style="height:100%; width:100%; display:flex; flex-direction:column; justify-content:center; align-items:center; padding:20px;">
            <div style="max-width:600px; width:100%; background-color:white; border-radius:10px; box-shadow:0 4px 20px rgba(0,0,0,0.1); padding:30px; text-align:center;">
                <div style="font-size:64px; margin-bottom:20px;">${fileIcon}</div>
                <h2 style="margin-bottom:16px; color:#333; font-size:24px; word-break:break-word;">${fileName}</h2>
                <div style="background-color:#f8f9fa; border-radius:6px; padding:12px; margin-bottom:24px; display:inline-block; color:#5f6368; font-size:14px;">
                    ${getFileTypeName(fileType)}
                </div>
                
                <div style="background-color:#e8f0fe; border-left:4px solid #4285f4; padding:16px; margin:24px 0; text-align:left; border-radius:4px;">
                    <p style="margin:0 0 12px 0; color:#1a73e8; font-weight:500;">关于文档预览</p>
                    <p style="margin:0; color:#333; line-height:1.5;">Word文档在插件内无法直接预览。您可以使用以下选项查看文档内容：</p>
                </div>
                
                <div style="display:flex; gap:16px; justify-content:center; margin-top:30px; flex-wrap:wrap;">
                    <button id="word-drive-view" class="ms-Button ms-Button--primary" style="min-width:150px;">
                        <span style="margin-right:8px;">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                                <path d="M19 19H5V5h7V3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2v-7h-2v7zM14 3v2h3.59l-9.83 9.83 1.41 1.41L19 6.41V10h2V3h-7z"/>
                            </svg>
                        </span>
                        在Drive中查看
                    </button>
                    <button id="word-download" class="ms-Button" style="min-width:150px;">
                        <span style="margin-right:8px;">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                                <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
                            </svg>
                        </span>
                        下载文档
                    </button>
                </div>
                
                <div style="margin-top:30px; padding-top:20px; border-top:1px solid #e8eaed; text-align:left;">
                    <h3 style="font-size:16px; margin-bottom:12px; color:#333;">需要在Word中插入内容？</h3>
                    <p style="margin:0 0 10px 0; color:#5f6368; font-size:14px; line-height:1.5;">
                        您可以下载文档后复制所需内容，或直接点击下方按钮尝试插入文本内容：
                    </p>
                    <button id="insert-doc-name" class="ms-Button" style="margin-top:8px;">
                        <span style="margin-right:8px;">➕</span>
                        插入文档名称和链接
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // 绑定按钮事件
    document.getElementById('word-drive-view').addEventListener('click', function() {
        window.open(file.webViewLink, '_blank');
    });
    
    document.getElementById('word-download').addEventListener('click', function() {
        downloadFile(file.id, file.name);
    });
    
    document.getElementById('insert-doc-name').addEventListener('click', function() {
        // 创建一个包含文档名称和链接的简单文本
        const insertText = `${file.name}\n\n查看原始文档: ${file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`}`;
        
        // 插入到Word文档
        Office.context.document.setSelectedDataAsync(
            insertText,
            { coercionType: Office.CoercionType.Text },
            function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    alert(`插入失败: ${result.error.message}`);
                } else {
                    alert('已成功插入文档名称和链接');
                }
            }
        );
    });
}

export { extractDocumentText, renderSimpleGoogleDocPreview, renderSimpleWordPreview };