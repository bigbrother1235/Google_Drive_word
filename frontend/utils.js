// utils.js - 通用工具函数

// 日志函数：安全地记录日志
function safeLog(...args) {
    try {
        const timestamp = new Date().toISOString().split('T')[1].split('.')[0]; // 简单时间戳 HH:MM:SS
        console.log(`[${timestamp}]`, ...args);
    } catch (e) {
        // 兼容性处理
    }
}

// 更新状态信息
function updateStatus(message, isError = false) {
    const statusBox = document.getElementById('auth-status');
    if (statusBox) {
        statusBox.textContent = message;
        statusBox.style.color = isError ? 'red' : 'black';
        statusBox.style.backgroundColor = isError ? '#ffeeee' : '#f5f5f5';
    }
}

// 错误处理函数
function handleError(error, isAuthError = false) {
    safeLog('Error:', error);
    
    if (isAuthError || (error && error.message && (
        error.message.includes('401') || 
        error.message.toLowerCase().includes('unauthorized') ||
        error.message.toLowerCase().includes('unauthenticated')
    ))) {
        // 处理授权错误
        showTokenRefreshModal();
        updateStatus('授权已过期，请重新授权', true);
    } else {
        updateStatus(`错误: ${error.message || error.toString()}`, true);
    }
}

// 复制文本的后备方法
function copyTextFallback(text) {
    try {
        // 创建临时文本区域
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-9999px';
        textArea.style.top = '-9999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        
        const success = document.execCommand('copy');
        document.body.removeChild(textArea);
        
        if (success) {
            alert('文本已复制到剪贴板');
        } else {
            alert('复制失败，请手动选择复制');
        }
    } catch (err) {
        console.error('后备复制方法失败:', err);
        alert('复制失败，请手动选择复制');
    }
}

// HTML安全清理函数
function sanitizeHtml(html) {
    // 使用简单的方法移除潜在危险标签
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    
    // 移除所有脚本标签
    const scripts = tempDiv.getElementsByTagName('script');
    while (scripts[0]) scripts[0].parentNode.removeChild(scripts[0]);
    
    // 移除所有iframe标签
    const iframes = tempDiv.getElementsByTagName('iframe');
    while (iframes[0]) iframes[0].parentNode.removeChild(iframes[0]);
    
    // 移除所有事件属性
    const allElements = tempDiv.getElementsByTagName('*');
    for (let i = 0; i < allElements.length; i++) {
        const element = allElements[i];
        const attributes = element.attributes;
        for (let j = attributes.length - 1; j >= 0; j--) {
            const name = attributes[j].name;
            if (name.startsWith('on') || name === 'href' && attributes[j].value.startsWith('javascript:')) {
                element.removeAttribute(name);
            }
        }
    }
    
    return tempDiv.innerHTML;
}

// 获取人类可读的文件类型名称
function getFileTypeName(mimeType) {
    const typeMap = {
        'application/vnd.google-apps.folder': '文件夹',
        'application/vnd.google-apps.document': 'Google 文档',
        'application/vnd.google-apps.spreadsheet': 'Google 表格',
        'application/vnd.google-apps.presentation': 'Google 幻灯片',
        'application/pdf': 'PDF 文件',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word 文档',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel 表格',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint 演示文稿',
        'application/msword': 'Word 文档',
        'application/vnd.ms-excel': 'Excel 表格',
        'application/vnd.ms-powerpoint': 'PowerPoint 演示文稿',
        'image/jpeg': 'JPEG 图片',
        'image/png': 'PNG 图片',
        'image/gif': 'GIF 图片',
        'image/svg+xml': 'SVG 图片',
        'text/plain': '文本文件',
        'text/html': 'HTML 文件',
        'application/json': 'JSON 文件',
        'application/zip': 'ZIP 压缩文件',
        'video/mp4': 'MP4 视频',
        'audio/mpeg': 'MP3 音频'
    };
    
    return typeMap[mimeType] || mimeType;
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 B';
    
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

export { safeLog, updateStatus, handleError, copyTextFallback, sanitizeHtml, getFileTypeName, formatFileSize };