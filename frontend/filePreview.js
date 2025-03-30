// filePreview.js - 处理文件预览相关功能

import { safeLog, updateStatus, handleError, sanitizeHtml, formatFileSize, getFileTypeName } from './utils.js';
import { extractDocumentText, renderSimpleGoogleDocPreview, renderSimpleWordPreview } from './documentHandler.js';
import { renderSimplePdfPreview } from './pdfHandler.js';

// 预览文件入口函数
function previewFile(file) {
    // 保存当前活动文件
    activeFile = file;
    
    const modal = document.getElementById('file-preview-modal');
    const previewFileName = document.getElementById('preview-file-name');
    const fileInfo = document.getElementById('file-info');
    const filePreviewContent = document.getElementById('file-preview-content');
    const openInDriveBtn = document.getElementById('open-in-drive');
    const downloadFileBtn = document.getElementById('download-file');
    
    // 显示模态框
    modal.style.display = 'block';
    
    // 设置文件名
    previewFileName.textContent = file.name;
    
    // 清空预览内容区域，显示加载中
    fileInfo.innerHTML = '<div><span class="loading"></span> 正在加载文件信息...</div>';
    filePreviewContent.innerHTML = '<div style="display:flex; justify-content:center; align-items:center; height:300px;"><span class="loading"></span><div style="margin-left:10px;">正在加载预览...</div></div>';
    
    // 配置Drive打开按钮
    if (file.webViewLink) {
        openInDriveBtn.onclick = function() {
            window.open(file.webViewLink, '_blank');
        };
        openInDriveBtn.style.display = 'inline-block';
    } else {
        openInDriveBtn.style.display = 'none';
    }
    
    // 配置下载按钮
    if (file.mimeType !== 'application/vnd.google-apps.folder') {
        downloadFileBtn.onclick = function() {
            downloadFile(file.id, file.name);
        };
        downloadFileBtn.style.display = 'inline-block';
    } else {
        downloadFileBtn.style.display = 'none';
    }
    
    // 获取文件详情
    if (!tokenInfo) {
        fileInfo.innerHTML = '<div class="error-message">未授权，无法获取文件详情</div>';
        return;
    }
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        scopes: tokenInfo.scopes,
        file_id: file.id
    };
    
    // 发起请求获取文件详情
    fetch(`${API_BASE_URL}/api/file-details`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            if (response.status === 401) {
                throw new Error('授权已过期，需要重新登录');
            }
            return response.json().then(errorData => {
                throw new Error(`服务器错误: ${errorData.error || '未知错误'}`);
            });
        }
        return response.json();
    })
    .then(data => {
        safeLog('获取文件详情成功', data);
        
        // 更新文件信息
        const fileDetails = data.file;
        const fileSize = fileDetails.size ? formatFileSize(parseInt(fileDetails.size)) : '未知';
        const modifiedDate = fileDetails.modifiedTime ? new Date(fileDetails.modifiedTime).toLocaleString() : '未知';
        const createdDate = fileDetails.createdTime ? new Date(fileDetails.createdTime).toLocaleString() : '未知';
        
        fileInfo.innerHTML = `
            <div><strong>名称:</strong> ${fileDetails.name}</div>
            <div><strong>类型:</strong> ${getFileTypeName(fileDetails.mimeType)}</div>
            <div><strong>大小:</strong> ${fileSize}</div>
            <div><strong>修改时间:</strong> ${modifiedDate}</div>
            <div><strong>创建时间:</strong> ${createdDate}</div>
        `;
        
        // 根据文件类型简单预览
        simplifiedPreview(fileDetails, data.content);
    })
    .catch(error => {
        fileInfo.innerHTML = `<div class="error-message">获取文件详情失败: ${error.message}</div>`;
        filePreviewContent.innerHTML = renderErrorState(file);
        
        if (error.message.includes('401')) {
            handleError(error, true);
        }
    });
}

// 简化的文件预览
function simplifiedPreview(file, content) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    // 清空类名
    filePreviewContent.className = '';
    
    // 根据文件类型选择预览方式
    switch(file.mimeType) {
        case 'application/pdf':
            renderSimplePdfPreview(file);
            break;
            
        case 'application/vnd.google-apps.document':
            renderSimpleGoogleDocPreview(file, content);
            break;
            
        case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        case 'application/msword':
            renderSimpleWordPreview(file);
            break;
            
        case 'image/jpeg':
        case 'image/png':
        case 'image/gif':
        case 'image/bmp':
        case 'image/webp':
        case 'image/svg+xml':
            renderSimpleImagePreview(file);
            break;
            
        case 'text/plain':
        case 'text/html':
        case 'text/css':
        case 'text/javascript':
        case 'application/json':
        case 'application/xml':
            renderSimpleTextPreview(file);
            break;
            
        default:
            renderGenericPreview(file);
            break;
    }
}

// 显示错误状态
function renderErrorState(file) {
    return `
        <div style="text-align:center; padding:40px 20px; background-color:white; border-radius:8px; margin:20px;">
            <div style="font-size:48px; margin-bottom:20px;">⚠️</div>
            <h3 style="margin-bottom:20px; color:#d93025;">无法预览文件</h3>
            <p style="color:#5f6368; margin-bottom:30px;">我们遇到了一些问题，无法加载此文件的预览。</p>
            <div style="display:flex; gap:16px; justify-content:center;">
                <button onclick="downloadFile('${file.id}', '${file.name}')" class="ms-Button ms-Button--primary">
                    下载文件
                </button>
                <button onclick="window.open('${file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`}', '_blank')" class="ms-Button">
                    在Drive中查看
                </button>
            </div>
        </div>
    `;
}

// 简化的图片预览 - 使用Google Drive API直接请求图片
function renderSimpleImagePreview(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    filePreviewContent.innerHTML = '<div style="display:flex; justify-content:center; align-items:center; height:100%; width:100%;"><div class="loading"></div><p>正在加载图片...</p></div>';
    
    // 直接创建一个带有认证令牌的图片请求URL
    const imgUrl = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
    
    // 创建一个新的XMLHttpRequest来获取图片数据
    const xhr = new XMLHttpRequest();
    xhr.open('GET', imgUrl, true);
    xhr.setRequestHeader('Authorization', `Bearer ${tokenInfo.token}`);
    xhr.responseType = 'blob';
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            // 创建一个临时URL来显示图片
            const blob = xhr.response;
            const objectUrl = URL.createObjectURL(blob);
            
            // 显示图片和控制按钮
            filePreviewContent.innerHTML = `
                <div style="height:100%; width:100%; display:flex; flex-direction:column; justify-content:center; align-items:center; background-color:#2d2d2d; padding:20px;">
                    <img src="${objectUrl}" id="preview-img" style="max-width:90%; max-height:80%; object-fit:contain;">
                    <div style="margin-top:20px;">
                        <button id="img-copy" class="ms-Button ms-Button--primary" style="margin-right:12px;">
                            复制图片
                        </button>
                        <button id="img-download" class="ms-Button">
                            下载图片
                        </button>
                    </div>
                </div>
            `;
            
            // 下载按钮
            document.getElementById('img-download').addEventListener('click', function() {
                downloadFile(file.id, file.name);
            });
            
            // 复制按钮 - 尝试复制到剪贴板
            document.getElementById('img-copy').addEventListener('click', function() {
                const img = document.getElementById('preview-img');
                copyImageToClipboard(img, blob);
            });
        } else {
            // 如果请求失败，显示错误信息并提供备选方案
            filePreviewContent.innerHTML = `
                <div class="error-message" style="text-align:center">
                    <p>无法加载图片 (错误: ${xhr.status})</p>
                    <div style="margin-top:20px;">
                        <button id="img-alt-download" class="ms-Button ms-Button--primary" style="margin-right:12px;">
                            下载图片
                        </button>
                        <button id="img-drive-view" class="ms-Button">
                            在Drive中查看
                        </button>
                    </div>
                </div>
            `;
            
            // 下载按钮
            document.getElementById('img-alt-download').addEventListener('click', function() {
                downloadFile(file.id, file.name);
            });
            
            // 在Drive中查看按钮
            document.getElementById('img-drive-view').addEventListener('click', function() {
                window.open(file.webViewLink, '_blank');
            });
        }
    };
    
    xhr.onerror = function() {
        // 网络错误
        filePreviewContent.innerHTML = `
            <div class="error-message" style="text-align:center">
                <p>网络错误，无法加载图片</p>
                <div style="margin-top:20px;">
                    <button onclick="window.open('${file.webViewLink}', '_blank')" class="ms-Button ms-Button--primary">
                        在Drive中查看
                    </button>
                </div>
            </div>
        `;
    };
    
    xhr.send();
}

// 尝试复制图片到剪贴板
function copyImageToClipboard(imgElement, blob) {
    // 尝试使用现代剪贴板API
    if (navigator.clipboard && navigator.clipboard.write) {
        try {
            const item = new ClipboardItem({ 'image/png': blob });
            navigator.clipboard.write([item])
                .then(() => alert('图片已复制到剪贴板'))
                .catch(err => {
                    console.error('复制图片失败:', err);
                    alert('无法复制图片，请尝试使用"插入内容"按钮');
                });
        } catch (e) {
            console.error('复制到剪贴板失败:', e);
            alert('您的浏览器不支持图片复制，请尝试使用"插入内容"按钮');
        }
    } else {
        alert('您的浏览器不支持图片复制，请尝试使用"插入内容"按钮');
    }
}

// 简化的文本文件预览
function renderSimpleTextPreview(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    filePreviewContent.innerHTML = '<div style="display:flex; justify-content:center; align-items:center; height:100%; width:100%;"><div class="loading"></div><p>正在加载文本内容...</p></div>';
    
    // 创建一个带认证的请求
    const textUrl = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media`;
    
    fetch(textUrl, {
        headers: {
            'Authorization': `Bearer ${tokenInfo.token}`
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP错误: ${response.status}`);
        }
        return response.text();
    })
    .then(text => {
        // 显示文本内容
        filePreviewContent.innerHTML = `
            <div style="height:100%; width:100%; display:flex; flex-direction:column; padding:20px;">
                <div style="display:flex; justify-content:flex-end; margin-bottom:16px;">
                    <button id="copy-text-content" class="ms-Button ms-Button--primary">复制文本</button>
                </div>
                <div style="flex:1; background-color:white; padding:20px; border-radius:8px; box-shadow:0 2px 10px rgba(0,0,0,0.1); overflow:auto;">
                    <pre style="white-space:pre-wrap; font-family:Monaco, Consolas, monospace; font-size:14px; line-height:1.5;">${text}</pre>
                </div>
            </div>
        `;
        
        // 添加复制按钮功能
        document.getElementById('copy-text-content').addEventListener('click', function() {
            navigator.clipboard.writeText(text)
                .then(() => alert('文本已复制到剪贴板'))
                .catch(err => {
                    console.error('复制失败:', err);
                    copyTextFallback(text);
                });
        });
    })
    .catch(error => {
        // 显示错误并提供备选方案
        filePreviewContent.innerHTML = `
            <div class="error-message" style="text-align:center">
                <p>无法加载文本内容: ${error.message}</p>
                <div style="margin-top:20px;">
                    <button id="text-download" class="ms-Button ms-Button--primary" style="margin-right:12px;">
                        下载文件
                    </button>
                    <button id="text-drive-view" class="ms-Button">
                        在Drive中查看
                    </button>
                </div>
            </div>
        `;
        
        // 下载按钮
        document.getElementById('text-download').addEventListener('click', function() {
            downloadFile(file.id, file.name);
        });
        
        // 在Drive中查看按钮
        document.getElementById('text-drive-view').addEventListener('click', function() {
            window.open(file.webViewLink, '_blank');
        });
    });
}

// 渲染通用预览
function renderGenericPreview(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    const fileIcon = getFileIcon(file.mimeType);
    const fileTypeName = getFileTypeName(file.mimeType);
    
    filePreviewContent.innerHTML = `
        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:40px 20px; background-color:white; border-radius:8px; margin:20px;">
            <div style="font-size:64px; margin-bottom:20px;">${fileIcon}</div>
            <h3 style="margin-bottom:20px; text-align:center;">${file.name}</h3>
            <p style="color:#5f6368; margin-bottom:30px; text-align:center;">此 ${fileTypeName} 文件无法在浏览器中预览</p>
            <div style="display:flex; gap:16px;">
                <button id="generic-download" class="ms-Button ms-Button--primary">
                    下载文件
                </button>
                <button id="generic-drive-view" class="ms-Button">
                    在Drive中查看
                </button>
            </div>
        </div>
    `;
    
    // 下载按钮
    document.getElementById('generic-download').addEventListener('click', function() {
        downloadFile(file.id, file.name);
    });
    
    // 在Drive中查看按钮
    document.getElementById('generic-drive-view').addEventListener('click', function() {
        window.open(file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`, '_blank');
    });
}

// 下载文件
function downloadFile(fileId, fileName) {
    if (!tokenInfo) {
        alert('未授权，请先登录');
        return;
    }
    
    // 构建下载URL
    const downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`;
    
    // 创建一个隐藏的a标签用于下载
    const a = document.createElement('a');
    a.style.display = 'none';
    document.body.appendChild(a);
    
    // 显示下载中提示
    const loadingDiv = document.createElement('div');
    loadingDiv.className = 'download-loading';
    loadingDiv.style.position = 'fixed';
    loadingDiv.style.top = '50%';
    loadingDiv.style.left = '50%';
    loadingDiv.style.transform = 'translate(-50%, -50%)';
    loadingDiv.style.backgroundColor = 'rgba(255,255,255,0.9)';
    loadingDiv.style.padding = '20px';
    loadingDiv.style.borderRadius = '8px';
    loadingDiv.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
    loadingDiv.style.zIndex = '1000';
    loadingDiv.innerHTML = '<div class="loading"></div><div style="margin-top:10px;">正在下载文件...</div>';
    document.body.appendChild(loadingDiv);
    
    // 使用XMLHttpRequest获取文件
    const xhr = new XMLHttpRequest();
    xhr.open('GET', downloadUrl, true);
    xhr.setRequestHeader('Authorization', `Bearer ${tokenInfo.token}`);
    xhr.responseType = 'blob';
    
    xhr.onload = function() {
        if (xhr.status === 200) {
            const blob = xhr.response;
            const url = window.URL.createObjectURL(blob);
            a.href = url;
            a.download = fileName;
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(loadingDiv);
        } else {
            document.body.removeChild(loadingDiv);
            if (xhr.status === 401) {
                alert('授权已过期，请重新登录');
                showTokenRefreshModal();
            } else {
                alert(`下载失败: HTTP错误 ${xhr.status}`);
            }
        }
        document.body.removeChild(a);
    };
    
    xhr.onerror = function() {
        document.body.removeChild(loadingDiv);
        document.body.removeChild(a);
        alert('网络错误，下载失败');
    };
    
    xhr.send();
}

// 将选中的内容插入到Office文档中
function insertSelectedContent() {
    if (!activeFile) {
        alert('没有活动文件');
        return;
    }
    
    // 插入前显示加载状态
    const modal = document.getElementById('file-preview-modal');
    const loadingDiv = document.createElement('div');
    loadingDiv.className = 'insert-loading';
    loadingDiv.style.position = 'fixed';
    loadingDiv.style.top = '50%';
    loadingDiv.style.left = '50%';
    loadingDiv.style.transform = 'translate(-50%, -50%)';
    loadingDiv.style.backgroundColor = 'rgba(255,255,255,0.9)';
    loadingDiv.style.padding = '20px';
    loadingDiv.style.borderRadius = '8px';
    loadingDiv.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
    loadingDiv.style.zIndex = '2000';
    loadingDiv.innerHTML = '<div class="loading"></div><div style="margin-top:10px;">正在处理内容...</div>';
    document.body.appendChild(loadingDiv);
    
    // 根据文件类型处理
    switch (activeFile.mimeType) {
        case 'application/pdf':
            insertPdfContent();
            break;
            
        case 'application/vnd.google-apps.document':
            insertGoogleDocContent();
            break;
            
        case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        case 'application/msword':
            insertWordContent();
            break;
            
        case 'image/jpeg':
        case 'image/png':
        case 'image/gif':
        case 'image/bmp':
        case 'image/webp':
        case 'image/svg+xml':
            insertImageContent();
            break;
            
        case 'text/plain':
        case 'text/html':
        case 'text/css':
        case 'text/javascript':
        case 'application/json':
        case 'application/xml':
            insertTextContent();
            break;
            
        default:
            // 不支持的类型
            document.body.removeChild(loadingDiv);
            alert(`不支持将 ${getFileTypeName(activeFile.mimeType)} 类型的内容插入到文档中。请尝试下载后手动添加。`);
            break;
    }
}

// 插入PDF内容
function insertPdfContent() {
    // 检查是否已经提取的PDF文本
    const extractedText = document.querySelector('#pdf-text-content pre')?.textContent;
    
    if (extractedText) {
        // 已提取的文本可直接插入
        insertTextToDocument(extractedText);
    } else {
        // 提取PDF文本并插入
        const requestBody = {
            token: tokenInfo.token,
            refresh_token: tokenInfo.refresh_token,
            token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
            client_id: tokenInfo.client_id,
            file_id: activeFile.id
        };
        
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
                insertTextToDocument(data.text);
            } else {
                throw new Error('未能提取文本内容');
            }
        })
        .catch(error => {
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            alert(`无法提取PDF内容: ${error.message}`);
        });
    }
}

// 插入Google文档内容
function insertGoogleDocContent() {
    // 检查是否已显示了文档内容
    const docContent = document.querySelector('.doc-content')?.innerText;
    const extractedText = document.querySelector('#doc-text-container pre')?.textContent;
    
    if (docContent) {
        // 已显示的文档内容可直接插入
        insertTextToDocument(docContent);
    } else if (extractedText) {
        // 已提取的文本可直接插入
        insertTextToDocument(extractedText);
    } else {
        // 提取Google文档内容并插入
        const requestBody = {
            token: tokenInfo.token,
            refresh_token: tokenInfo.refresh_token,
            token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
            client_id: tokenInfo.client_id,
            file_id: activeFile.id,
            mime_type: activeFile.mimeType
        };
        
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
                insertTextToDocument(data.text);
            } else {
                throw new Error('未能提取文本内容');
            }
        })
        .catch(error => {
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            alert(`无法提取文档内容: ${error.message}`);
        });
    }
}

// 插入Word文档内容
function insertWordContent() {
    // 检查是否已提取的文本
    const extractedText = document.querySelector('#word-text-container pre')?.textContent;
    
    if (extractedText) {
        // 已提取的文本可直接插入
        insertTextToDocument(extractedText);
    } else {
        // 提取Word文档内容并插入
        const requestBody = {
            token: tokenInfo.token,
            refresh_token: tokenInfo.refresh_token,
            token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
            client_id: tokenInfo.client_id,
            file_id: activeFile.id,
            mime_type: activeFile.mimeType
        };
        
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
                insertTextToDocument(data.text);
            } else {
                throw new Error('未能提取文本内容');
            }
        })
        .catch(error => {
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            alert(`无法提取Word文档内容: ${error.message}`);
        });
    }
}

// 插入图片内容
function insertImageContent() {
    const imgElement = document.querySelector('#preview-img');
    
    if (imgElement && imgElement.src) {
        // 如果已加载图片，直接插入
        Office.context.document.setSelectedDataAsync(
            imgElement.src,
            { coercionType: Office.CoercionType.Image },
            function(result) {
                const loadingDiv = document.querySelector('.insert-loading');
                if (loadingDiv) document.body.removeChild(loadingDiv);
                
                if (result.status === Office.AsyncResultStatus.Failed) {
                    alert(`插入图片失败: ${result.error.message}`);
                } else {
                    // 成功插入，关闭预览
                    document.getElementById('file-preview-modal').style.display = 'none';
                }
            }
        );
    } else {
        // 需要获取图片
        const imgUrl = `https://www.googleapis.com/drive/v3/files/${activeFile.id}?alt=media`;
        
        const xhr = new XMLHttpRequest();
        xhr.open('GET', imgUrl, true);
        xhr.setRequestHeader('Authorization', `Bearer ${tokenInfo.token}`);
        xhr.responseType = 'blob';
        
        xhr.onload = function() {
            if (xhr.status === 200) {
                const blob = xhr.response;
                const reader = new FileReader();
                
                reader.onloadend = function() {
                    const base64data = reader.result;
                    
                    Office.context.document.setSelectedDataAsync(
                        base64data,
                        { coercionType: Office.CoercionType.Image },
                        function(result) {
                            const loadingDiv = document.querySelector('.insert-loading');
                            if (loadingDiv) document.body.removeChild(loadingDiv);
                            
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                alert(`插入图片失败: ${result.error.message}`);
                            } else {
                                // 成功插入，关闭预览
                                document.getElementById('file-preview-modal').style.display = 'none';
                            }
                        }
                    );
                };
                
                reader.readAsDataURL(blob);
            } else {
                const loadingDiv = document.querySelector('.insert-loading');
                if (loadingDiv) document.body.removeChild(loadingDiv);
                
                alert(`获取图片数据失败: HTTP错误 ${xhr.status}`);
            }
        };
        
        xhr.onerror = function() {
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            alert('网络错误，获取图片数据失败');
        };
        
        xhr.send();
    }
}

// 插入文本内容
function insertTextContent() {
    // 检查是否已加载文本
    const preElement = document.querySelector('#file-preview-content pre');
    
    if (preElement && preElement.textContent) {
        // 已加载的文本可直接插入
        insertTextToDocument(preElement.textContent);
    } else {
        // 需要获取文本
        const textUrl = `https://www.googleapis.com/drive/v3/files/${activeFile.id}?alt=media`;
        
        fetch(textUrl, {
            headers: {
                'Authorization': `Bearer ${tokenInfo.token}`
            }
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP错误: ${response.status}`);
            }
            return response.text();
        })
        .then(text => {
            insertTextToDocument(text);
        })
        .catch(error => {
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            alert(`获取文本内容失败: ${error.message}`);
        });
    }
}

// 插入文本到文档
function insertTextToDocument(text) {
    Office.context.document.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        function(result) {
            // 移除加载状态
            const loadingDiv = document.querySelector('.insert-loading');
            if (loadingDiv) document.body.removeChild(loadingDiv);
            
            if (result.status === Office.AsyncResultStatus.Failed) {
                alert(`插入文本失败: ${result.error.message}`);
            } else {
                // 关闭预览
                document.getElementById('file-preview-modal').style.display = 'none';
            }
        }
    );
}

export { previewFile, insertSelectedContent, downloadFile, simplifiedPreview, renderSimpleImagePreview, renderSimpleTextPreview, renderGenericPreview, renderErrorState };
        // 提