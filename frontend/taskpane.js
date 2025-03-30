// 后端API基础URL
const API_BASE_URL = 'https://localhost:5001';

// 当前的令牌信息
let tokenInfo = null;

// 当前的导航路径信息
const breadcrumbPath = [
    { id: 'root', name: '根目录' }
];

// 记录MIME类型对应的图标
const mimeTypeIcons = {
    'folder': '📁',
    'document': '📄',
    'spreadsheet': '📊',
    'presentation': '📽️',
    'image': '🖼️',
    'audio': '🎵',
    'video': '🎬',
    'pdf': '📑',
    'archive': '📦',
    'word': '📝',
    'excel': '📈',
    'powerpoint': '📺',
    'unknown': '📎'
};

// 当前活动文件
let activeFile = null;

// 添加全局变量 - 搜索支持
let isSearchMode = false;
let currentSearchTerm = '';

// 添加全局变量 - 令牌刷新
let tokenRefreshTimeout = null;
const TOKEN_REFRESH_INTERVAL = 45 * 60 * 1000; // 45分钟
const TOKEN_REFRESH_WARNING = 50 * 60 * 1000; // 50分钟

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

// 显示令牌刷新提示
function showTokenRefreshModal() {
    const modal = document.getElementById('token-refresh-modal');
    if (modal) {
        modal.style.display = 'block';
        
        // 添加刷新按钮事件
        const refreshButton = document.getElementById('refresh-token-button');
        if (refreshButton) {
            refreshButton.onclick = function() {
                modal.style.display = 'none';
                handleAuthClick();
            };
        }
    }
}

// 刷新令牌
function setupTokenRefresh() {
    // 清除任何现有的超时
    if (tokenRefreshTimeout) {
        clearTimeout(tokenRefreshTimeout);
    }
    
    // 设置新的超时 - 在令牌过期前显示提醒
    tokenRefreshTimeout = setTimeout(function() {
        showTokenRefreshModal();
    }, TOKEN_REFRESH_WARNING);
    
    safeLog('令牌刷新定时器已设置');
}

// 更新UI状态
function updateUIState(isAuthorized) {
    const authSection = document.getElementById('auth-section');
    const browserSection = document.getElementById('browser-section');
    const authButton = document.getElementById('auth-button');
    
    if (isAuthorized) {
        // 显示已授权状态
        authButton.textContent = '重新授权';
        browserSection.style.display = 'block';
        updateStatus('已成功授权Google Drive');
        
        // 设置令牌刷新提醒
        setupTokenRefresh();
        
        // 加载根目录内容
        loadFolderContents('root');
    } else {
        // 显示未授权状态
        authButton.textContent = '授权 Google Drive';
        browserSection.style.display = 'none';
        updateStatus('请点击上方按钮授权访问Google Drive');
    }
}

// 处理授权按钮点击
function handleAuthClick() {
    updateStatus('正在获取授权URL...');
    
    fetch(`${API_BASE_URL}/api/get-auth-url`, {
        method: 'GET',
        headers: {
            'Accept': 'application/json'
        }
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP错误! 状态: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        safeLog('获取授权URL成功', data);
        
        // 修改：不再使用localStorage存储state
        // 而是将整个授权URL显示给用户
        const authUrl = data.auth_url;
        
        // 更新状态显示授权链接
        const statusBox = document.getElementById('auth-status');
        statusBox.innerHTML = `
            <div>请点击下面的链接完成授权:</div>
            <div style="margin:10px 0;">
                <a href="${authUrl}" target="_blank" class="auth-link">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 48 48" style="vertical-align: middle; margin-right: 8px;">
                        <path fill="#FFC107" d="M43.611 20.083H42V20H24v8h11.303c-1.649 4.657-6.08 8-11.303 8c-6.627 0-12-5.373-12-12s5.373-12 12-12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4C12.955 4 4 12.955 4 24s8.955 20 20 20s20-8.955 20-20c0-1.341-.138-2.65-.389-3.917z"/>
                        <path fill="#FF3D00" d="m6.306 14.691l6.571 4.819C14.655 15.108 18.961 12 24 12c3.059 0 5.842 1.154 7.961 3.039l5.657-5.657C34.046 6.053 29.268 4 24 4C16.318 4 9.656 8.337 6.306 14.691z"/>
                        <path fill="#4CAF50" d="M24 44c5.166 0 9.86-1.977 13.409-5.192l-6.19-5.238A11.91 11.91 0 0 1 24 36c-5.202 0-9.619-3.317-11.283-7.946l-6.522 5.025C9.505 39.556 16.227 44 24 44z"/>
                        <path fill="#1976D2" d="M43.611 20.083H42V20H24v8h11.303a12.04 12.04 0 0 1-4.087 5.571l.003-.002l6.19 5.238C36.971 39.205 44 34 44 24c0-1.341-.138-2.65-.389-3.917z"/>
                    </svg>
                    授权Google Drive
                </a>
            </div>
            <div style="margin-top:16px; color:#666; background-color:#f8f9fa; padding:12px; border-radius:6px; border-left:4px solid #4285F4;">
                授权完成后，请复制浏览器地址栏中的完整URL，然后粘贴到下面的输入框中:
            </div>
            <div style="margin-top:16px;">
                <input type="text" id="callback-url" placeholder="粘贴授权后的URL..." style="width:100%; padding:10px; border-radius:6px; border:1px solid #dadce0;">
                <button id="process-callback" style="margin-top:12px; background-color:#4285F4; color:white; padding:10px 16px; border-radius:6px; font-weight:500;">处理授权回调</button>
            </div>
        `;
        
        // 添加回调处理按钮事件
        document.getElementById('process-callback').onclick = function() {
            const callbackUrl = document.getElementById('callback-url').value;
            if (!callbackUrl) {
                updateStatus('请粘贴授权后的URL', true);
                return;
            }
            
            try {
                // 从粘贴的URL中提取code和state参数
                const url = new URL(callbackUrl);
                const code = url.searchParams.get('code');
                const state = url.searchParams.get('state');
                
                if (!code || !state) {
                    updateStatus('无效的回调URL，未找到code或state参数', true);
                    return;
                }
                
                // 处理授权码
                processAuthCode(code, state);
            } catch (e) {
                updateStatus('无效的URL格式', true);
            }
        };
    })
    .catch(handleError);
}

// 处理授权码的函数
function processAuthCode(code, state) {
    updateStatus('正在处理授权...');
    
    // 发送授权码到后端交换令牌
    fetch(`${API_BASE_URL}/api/auth-callback`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ code, state })
    })
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP错误! 状态: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        safeLog('授权成功', data);
        
        // 保存令牌信息
        tokenInfo = data;
        localStorage.setItem('tokenInfo', JSON.stringify(tokenInfo));
        
        updateStatus('授权成功! 正在加载文件...');
        updateUIState(true);
    })
    .catch(handleError);
}

// 根据mime类型获取图标
function getFileIcon(mimeType) {
    if (mimeType === 'application/vnd.google-apps.folder') {
        return mimeTypeIcons.folder;
    } else if (mimeType.includes('image/')) {
        return mimeTypeIcons.image;
    } else if (mimeType.includes('audio/')) {
        return mimeTypeIcons.audio;
    } else if (mimeType.includes('video/')) {
        return mimeTypeIcons.video;
    } else if (mimeType === 'application/pdf') {
        return mimeTypeIcons.pdf;
    } else if (mimeType === 'application/vnd.google-apps.document') {
        return mimeTypeIcons.document;
    } else if (mimeType === 'application/vnd.google-apps.spreadsheet') {
        return mimeTypeIcons.spreadsheet;
    } else if (mimeType === 'application/vnd.google-apps.presentation') {
        return mimeTypeIcons.presentation;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        return mimeTypeIcons.word;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        return mimeTypeIcons.excel;
    } else if (mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
        return mimeTypeIcons.powerpoint;
    } else if (mimeType.includes('zip') || mimeType.includes('compressed') || mimeType.includes('archive')) {
        return mimeTypeIcons.archive;
    } else {
        return mimeTypeIcons.unknown;
    }
}

// 获取文件类型的CSS类
function getFileClass(mimeType) {
    if (mimeType === 'application/vnd.google-apps.folder') {
        return 'folder';
    } else if (mimeType.includes('image/')) {
        return 'image';
    } else if (
        mimeType === 'application/vnd.google-apps.document' || 
        mimeType === 'application/vnd.google-apps.spreadsheet' || 
        mimeType === 'application/vnd.google-apps.presentation' ||
        mimeType === 'application/pdf' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    ) {
        return 'document';
    } else {
        return 'other';
    }
}

// 加载文件夹内容
function loadFolderContents(folderId) {
    if (!tokenInfo) {
        handleError(new Error('未授权，请先完成Google授权'));
        return;
    }

    updateStatus('正在加载文件夹内容...');
    
    // 显示加载指示器
    const filesContainer = document.getElementById('files-container');
    filesContainer.innerHTML = '<div class="empty-message"><span class="loading"></span> 正在加载...</div>';
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        scopes: tokenInfo.scopes,
        folder_id: folderId
    };
    
    fetch(`${API_BASE_URL}/api/list-folder-contents`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(`服务器错误: ${errorData.error || '未知错误'}`);
            });
        }
        return response.json();
    })
    .then(data => {
        safeLog('获取文件夹内容成功', data);
        renderFolderContents(data.files || []);
        updateStatus(`已加载 ${data.files ? data.files.length : 0} 个项目`);
    })
    .catch(error => {
        if (error.message.includes('401')) {
            handleError(error, true);
        } else {
            handleError(error);
        }
    });
}

// 渲染文件夹内容
function renderFolderContents(files) {
    const container = document.getElementById('files-container');
    container.innerHTML = '';

    if (files.length === 0) {
        container.innerHTML = '<div class="empty-message">此文件夹为空</div>';
        return;
    }

    // 首先显示文件夹，然后显示文件
    const folders = files.filter(file => file.mimeType === 'application/vnd.google-apps.folder');
    const otherFiles = files.filter(file => file.mimeType !== 'application/vnd.google-apps.folder');
    
    // 排序：按名称升序
    folders.sort((a, b) => a.name.localeCompare(b.name));
    otherFiles.sort((a, b) => a.name.localeCompare(b.name));
    
    // 合并排序后的文件夹和文件
    const sortedFiles = [...folders, ...otherFiles];
    
    sortedFiles.forEach(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const fileIcon = getFileIcon(file.mimeType);
        const fileClass = getFileClass(file.mimeType);
        
        const fileItem = document.createElement('div');
        fileItem.className = `file-item ${fileClass}`;
        fileItem.setAttribute('data-id', file.id);
        fileItem.setAttribute('data-name', file.name);
        fileItem.setAttribute('data-mime-type', file.mimeType);
        if (file.webViewLink) {
            fileItem.setAttribute('data-web-link', file.webViewLink);
        }
        
        fileItem.innerHTML = `
            <div class="file-icon">${fileIcon}</div>
            <div class="file-name" title="${file.name}">${file.name}</div>
        `;
        
        if (isFolder) {
            fileItem.onclick = function() {
                navigateToFolder(file.id, file.name);
            };
        } else {
            fileItem.onclick = function() {
                previewFile(file);
            };
        }
        
        container.appendChild(fileItem);
    });
}

// 设置搜索UI
function setupSearchUI() {
    const searchButton = document.getElementById('search-button');
    const searchInput = document.getElementById('search-input');
    
    if (searchButton && searchInput) {
        // 搜索按钮点击
        searchButton.onclick = function() {
            const searchTerm = searchInput.value.trim();
            if (searchTerm) {
                searchFiles(searchTerm);
            }
        };
        
        // 搜索框回车键
        searchInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                const searchTerm = searchInput.value.trim();
                if (searchTerm) {
                    searchFiles(searchTerm);
                }
            }
        });
    }
}

// 搜索文件功能
function searchFiles(searchTerm) {
    if (!tokenInfo) {
        handleError(new Error('未授权，请先完成Google授权'));
        return;
    }
    
    // 更新UI为搜索模式
    isSearchMode = true;
    currentSearchTerm = searchTerm;
    
    // 显示加载状态
    const filesContainer = document.getElementById('files-container');
    filesContainer.innerHTML = '<div class="empty-message"><span class="loading"></span> 正在搜索...</div>';
    
    // 构建请求体
    const requestBody = {
        token: tokenInfo.token,
        refresh_token: tokenInfo.refresh_token,
        token_uri: tokenInfo.token_uri || 'https://oauth2.googleapis.com/token',
        client_id: tokenInfo.client_id,
        scopes: tokenInfo.scopes,
        search_term: searchTerm
    };
    
    // 发起API调用
    fetch(`${API_BASE_URL}/api/search-files`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(requestBody)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(errorData => {
                throw new Error(`服务器错误: ${errorData.error || '未知错误'}`);
            });
        }
        return response.json();
    })
    .then(data => {
        safeLog('搜索文件成功', data);
        
        // 更新文件浏览器标题
        const breadcrumbContainer = document.getElementById('breadcrumb-container');
        breadcrumbContainer.innerHTML = `
            <span class="breadcrumb-item" data-id="root">根目录</span>
            <span class="breadcrumb-item search-result">搜索结果: "${searchTerm}"</span>
            <span class="breadcrumb-item close-search" title="返回浏览">×</span>
        `;
        
        // 添加返回浏览的点击事件
        document.querySelector('.close-search').onclick = function() {
            isSearchMode = false;
            // 返回当前文件夹视图
            const currentFolder = breadcrumbPath[breadcrumbPath.length - 1];
            loadFolderContents(currentFolder.id);
            renderBreadcrumb();
        };
        
        // 渲染搜索结果
        renderSearchResults(data.files || []);
        updateStatus(`找到 ${data.files ? data.files.length : 0} 个匹配项`);
    })
    .catch(error => {
        if (error.message.includes('401')) {
            handleError(error, true);
        } else {
            handleError(error);
        }
    });
}

// 渲染搜索结果
function renderSearchResults(files) {
    const container = document.getElementById('files-container');
    container.innerHTML = '';

    if (files.length === 0) {
        container.innerHTML = '<div class="empty-message">没有找到匹配的文件</div>';
        return;
    }

    // 首先显示文件夹，然后显示文件
    const folders = files.filter(file => file.mimeType === 'application/vnd.google-apps.folder');
    const otherFiles = files.filter(file => file.mimeType !== 'application/vnd.google-apps.folder');
    
    // 排序：按名称升序
    folders.sort((a, b) => a.name.localeCompare(b.name));
    otherFiles.sort((a, b) => a.name.localeCompare(b.name));
    
    // 合并排序后的文件夹和文件
    const sortedFiles = [...folders, ...otherFiles];
    
    sortedFiles.forEach(file => {
        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
        const fileIcon = getFileIcon(file.mimeType);
        const fileClass = getFileClass(file.mimeType);
        
        const fileItem = document.createElement('div');
        fileItem.className = `file-item ${fileClass}`;
        fileItem.setAttribute('data-id', file.id);
        fileItem.setAttribute('data-name', file.name);
        fileItem.setAttribute('data-mime-type', file.mimeType);
        if (file.webViewLink) {
            fileItem.setAttribute('data-web-link', file.webViewLink);
        }
        
        fileItem.innerHTML = `
            <div class="file-icon">${fileIcon}</div>
            <div class="file-name" title="${file.name}">${file.name}</div>
        `;
        
        if (isFolder) {
            fileItem.onclick = function() {
                navigateToFolder(file.id, file.name);
                isSearchMode = false;
            };
        } else {
            fileItem.onclick = function() {
                previewFile(file);
            };
        }
        
        container.appendChild(fileItem);
    });
}

// 导航到指定文件夹
function navigateToFolder(folderId, folderName) {
    // 更新面包屑导航
    const currentIndex = breadcrumbPath.findIndex(item => item.id === folderId);
    
    if (currentIndex >= 0) {
        // 如果在路径中，则截断到该位置
        breadcrumbPath.splice(currentIndex + 1);
    } else {
        // 如果不在路径中，则添加到路径末尾
        breadcrumbPath.push({ id: folderId, name: folderName });
    }
    
    // 重新渲染面包屑
    renderBreadcrumb();
    
    // 加载文件夹内容
    loadFolderContents(folderId);
}

// 渲染面包屑导航
function renderBreadcrumb() {
    const container = document.getElementById('breadcrumb-container');
    container.innerHTML = '';
    
    breadcrumbPath.forEach((item, index) => {
        const breadcrumbItem = document.createElement('span');
        breadcrumbItem.className = 'breadcrumb-item';
        breadcrumbItem.textContent = item.name;
        breadcrumbItem.setAttribute('data-id', item.id);
        
        // 除最后一项外，其他项可点击
        if (index < breadcrumbPath.length - 1) {
            breadcrumbItem.onclick = function() {
                navigateToFolder(item.id, item.name);
            };
        }
        
        container.appendChild(breadcrumbItem);
    });
}

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

// 简化的PDF预览
// 直接在应用内预览PDF
// 直接在应用内预览PDF
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

// 渲染PDF页面函数
function renderPdfPage(pdf, pageNumber, scale) {
    const container = document.getElementById('pdf-viewer-container');
    
    // 清空容器内容，但保留之前可能存在的canvas以避免闪烁
    const oldCanvas = document.getElementById('pdf-canvas');
    const loadingIndicator = document.createElement('div');
    loadingIndicator.className = 'loading';
    loadingIndicator.style.position = 'absolute';
    loadingIndicator.style.top = '50%';
    loadingIndicator.style.left = '50%';
    loadingIndicator.style.transform = 'translate(-50%, -50%)';
    
    if (!oldCanvas) {
        container.innerHTML = '';
        container.appendChild(loadingIndicator);
    }
    
    // 更新页码显示
    document.getElementById('current-page').textContent = `第${pageNumber}页`;
    
    // 获取页面
    pdf.getPage(pageNumber).then(function(page) {
        // 计算页面大小
        const viewport = page.getViewport({scale: scale});
        
        // 准备canvas
        const canvas = oldCanvas || document.createElement('canvas');
        canvas.id = 'pdf-canvas';
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.3)';
        
        if (!oldCanvas) {
            container.innerHTML = '';
            container.appendChild(canvas);
        }
        
        // 渲染PDF页面到canvas
        const context = canvas.getContext('2d');
        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };
        
        page.render(renderContext).promise.then(function() {
            // 页面渲染完成
        }).catch(function(error) {
            container.innerHTML = `
                <div class="error-message">
                    <p>渲染PDF页面失败: ${error.message}</p>
                </div>
            `;
        });
    }).catch(function(error) {
        container.innerHTML = `
            <div class="error-message">
                <p>获取PDF页面失败: ${error.message}</p>
            </div>
        `;
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
// 直接在应用内预览Word文档
// 完全客户端的Word文档处理 - 不调用任何提取API
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
function tryWordPreviewMethods(file) {
    // 预览方法1：使用Office Online Viewer
    tryOfficeOnlineViewer(file);
}

// 方法1: 使用Office Online Viewer
function tryOfficeOnlineViewer(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    // 构建Office Online Viewer URL
    // 这里使用的是一个特定的格式，先从Google Drive获取文件并通过直接URL提供给Office Viewer
    const directFileUrl = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media&access_token=${encodeURIComponent(tokenInfo.token)}`;
    const encodedUrl = encodeURIComponent(directFileUrl);
    const officeViewerUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodedUrl}`;
    
    // 这里我们直接创建iframe预览，不使用fetch等方式预加载文件
    filePreviewContent.innerHTML = `
        <div style="height:100%; width:100%; position:relative;">
            <iframe src="${officeViewerUrl}" 
                    style="width:100%; height:100%; border:none;" 
                    allowfullscreen
                    id="word-preview-frame"></iframe>
            <div id="word-preview-loading" class="loading" 
                 style="position:absolute; top:50%; left:50%; transform:translate(-50%, -50%)"></div>
            
            <!-- 备用按钮区域，默认隐藏，在预览失败时显示 -->
            <div id="word-preview-fallback" style="position:absolute; top:0; left:0; width:100%; height:100%; 
                 background:rgba(255,255,255,0.9); display:none; flex-direction:column; 
                 justify-content:center; align-items:center; text-align:center;">
                <div style="font-size:48px; margin-bottom:20px;">📝</div>
                <h3 style="margin-bottom:20px;">Word 文档预览失败</h3>
                <p style="color:#5f6368; margin-bottom:30px; max-width:80%;">
                    Office预览服务无法加载此文档。您可以尝试以下选项：
                </p>
                <div style="display:flex; gap:16px; flex-wrap:wrap; justify-content:center;">
                    <button id="try-google-preview" class="ms-Button">
                        尝试Google预览
                    </button>
                    <button id="extract-word-text" class="ms-Button ms-Button--primary">
                        提取文档文本
                    </button>
                    <button id="word-download" class="ms-Button">
                        下载文档
                    </button>
                    <button id="word-drive-view" class="ms-Button">
                        在Drive中查看
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // 监听iframe加载
    const iframe = document.getElementById('word-preview-frame');
    const loadingIndicator = document.getElementById('word-preview-loading');
    
    // 这里设置一个超时，如果在一定时间内iframe没有加载完毕，就显示备用选项
    let previewTimeout = setTimeout(() => {
        if (loadingIndicator) {
            loadingIndicator.style.display = 'none';
        }
        
        // 显示备用按钮
        document.getElementById('word-preview-fallback').style.display = 'flex';
        
        // 绑定备用按钮事件
        document.getElementById('try-google-preview').addEventListener('click', function() {
            tryGoogleDocsViewer(file); 
        });
        
        document.getElementById('extract-word-text').addEventListener('click', function() {
            extractDocumentText(file.id, file.mimeType);
        });
        
        document.getElementById('word-download').addEventListener('click', function() {
            downloadFile(file.id, file.name);
        });
        
        document.getElementById('word-drive-view').addEventListener('click', function() {
            window.open(file.webViewLink, '_blank');
        });
    }, 15000); // 15秒超时
    
    // iframe加载完成
    iframe.onload = function() {
        clearTimeout(previewTimeout); // 清除超时
        loadingIndicator.style.display = 'none';
    };
}

// 方法2: 尝试使用Google Docs Viewer
function tryGoogleDocsViewer(file) {
    const filePreviewContent = document.getElementById('file-preview-content');
    
    // 显示加载指示器
    filePreviewContent.innerHTML = `
        <div style="height:100%; width:100%; display:flex; justify-content:center; align-items:center;">
            <div class="loading"></div>
            <p style="margin-left:10px;">正在尝试Google预览...</p>
        </div>
    `;
    
    // 直接文件URL
    const directFileUrl = `https://www.googleapis.com/drive/v3/files/${file.id}?alt=media&access_token=${encodeURIComponent(tokenInfo.token)}`;
    const encodedUrl = encodeURIComponent(directFileUrl);
    
    // Google Docs Viewer URL
    const googleViewerUrl = `https://docs.google.com/viewer?url=${encodedUrl}&embedded=true`;
    
    // 嵌入Google Docs预览
    filePreviewContent.innerHTML = `
        <div style="height:100%; width:100%; position:relative;">
            <iframe src="${googleViewerUrl}" 
                    style="width:100%; height:100%; border:none;" 
                    allowfullscreen
                    id="google-preview-frame"></iframe>
            <div id="google-preview-loading" class="loading" 
                 style="position:absolute; top:50%; left:50%; transform:translate(-50%, -50%)"></div>
                 
            <!-- 备用按钮区域，默认隐藏，在预览失败时显示 -->
            <div id="google-preview-fallback" style="position:absolute; top:0; left:0; width:100%; height:100%; 
                 background:rgba(255,255,255,0.9); display:none; flex-direction:column; 
                 justify-content:center; align-items:center; text-align:center;">
                <div style="font-size:48px; margin-bottom:20px;">📝</div>
                <h3 style="margin-bottom:20px;">无法预览Word文档</h3>
                <p style="color:#5f6368; margin-bottom:30px; max-width:80%;">
                    我们尝试了多种预览方式，但都未成功。您可以：
                </p>
                <div style="display:flex; gap:16px; flex-wrap:wrap; justify-content:center;">
                    <button id="extract-word-text-final" class="ms-Button ms-Button--primary">
                        提取文档文本
                    </button>
                    <button id="word-download-final" class="ms-Button">
                        下载文档
                    </button>
                    <button id="word-drive-view-final" class="ms-Button">
                        在Drive中查看
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // 监听iframe加载
    const iframe = document.getElementById('google-preview-frame');
    const loadingIndicator = document.getElementById('google-preview-loading');
    
    // 设置超时
    let previewTimeout = setTimeout(() => {
        if (loadingIndicator) {
            loadingIndicator.style.display = 'none';
        }
        
        // 显示最终备用选项
        document.getElementById('google-preview-fallback').style.display = 'flex';
        
        // 绑定按钮事件
        document.getElementById('extract-word-text-final').addEventListener('click', function() {
            extractDocumentText(file.id, file.mimeType);
        });
        
        document.getElementById('word-download-final').addEventListener('click', function() {
            downloadFile(file.id, file.name);
        });
        
        document.getElementById('word-drive-view-final').addEventListener('click', function() {
            window.open(file.webViewLink, '_blank');
        });
    }, 15000); // 15秒超时
    
    // iframe加载完成
    iframe.onload = function() {
        clearTimeout(previewTimeout);
        loadingIndicator.style.display = 'none';
    };
}

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

// Office加载完成后的初始化
Office.onReady(function(info) {
    safeLog('Office已加载', info);

    // 绑定授权按钮点击事件
    const authButton = document.getElementById('auth-button');
    if (authButton) {
        authButton.onclick = handleAuthClick;
    }
    
    // 绑定关闭预览模态框的事件
    const closeModalButtons = document.querySelectorAll('.close, #close-preview');
    closeModalButtons.forEach(button => {
        button.onclick = function() {
            document.getElementById('file-preview-modal').style.display = 'none';
            
            // 重置活动文件状态
            activeFile = null;
        };
    });
    
    // 设置全屏/取消全屏按钮
    const toggleFullscreenBtn = document.getElementById('toggle-fullscreen');
    if (toggleFullscreenBtn) {
        toggleFullscreenBtn.onclick = function() {
            const modalContent = document.querySelector('.modal-content');
            modalContent.classList.toggle('fullscreen');
        };
    }
    
    // 在Office中插入内容的按钮
    const insertContentBtn = document.getElementById('insert-content');
    if (insertContentBtn) {
        insertContentBtn.onclick = function() {
            insertSelectedContent();
        };
    }

    // 设置搜索UI
    setupSearchUI();

    // 检查是否有保存的令牌
    const savedTokenInfo = localStorage.getItem('tokenInfo');
    if (savedTokenInfo) {
        try {
            tokenInfo = JSON.parse(savedTokenInfo);
            // 验证令牌是否有效
            if (tokenInfo && tokenInfo.token) {
                updateUIState(true);
                
                // 设置令牌刷新定时器
                setupTokenRefresh();
            }
        } catch (e) {
            safeLog('解析保存的令牌时出错', e);
            localStorage.removeItem('tokenInfo');
        }
    } else {
        updateUIState(false);
    }
    
    // 点击模态框外部关闭模态框
    window.onclick = function(event) {
        const modal = document.getElementById('file-preview-modal');
        const tokenRefreshModal = document.getElementById('token-refresh-modal');
        
        if (event.target === modal) {
            modal.style.display = 'none';
            // 重置状态
            activeFile = null;
        }
        
        if (event.target === tokenRefreshModal) {
            tokenRefreshModal.style.display = 'none';
        }
    };
});

// 兜底的错误处理
window.addEventListener('error', function(event) {
    safeLog('全局错误捕获:', event.error);
});




